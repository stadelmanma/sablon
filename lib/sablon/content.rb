require 'open-uri'

module Sablon
  module Content
    class << self
      def wrap(value)
        case value
        when Sablon::Content
          value
        else
          if type = type_wrapping(value)
            type.new(value)
          else
            raise ArgumentError, "Could not find Sablon content type to wrap #{value.inspect}"
          end
        end
      end

      def make(type_id, *args)
        if types.key?(type_id)
          types[type_id].new(*args)
        else
          raise ArgumentError, "Could not find Sablon content type with id '#{type_id}'"
        end
      end

      def register(content_type)
        types[content_type.id] = content_type
      end

      def remove(content_type_or_id)
        types.delete_if {|k,v| k == content_type_or_id || v == content_type_or_id }
      end

      private
      def type_wrapping(value)
        types.values.reverse.detect { |type| type.wraps?(value) }
      end

      def types
        @types ||= {}
      end
    end

    # Handles simple text replacement of fields in the template
    class String < Struct.new(:string)
      include Sablon::Content
      def self.id; :string end
      def self.wraps?(value)
        value.respond_to?(:to_s)
      end

      def initialize(value)
        super value.to_s
      end

      def append_to(paragraph, display_node, env)
        string.scan(/[^\n]+|\n/).reverse.each do |part|
          if part == "\n"
            display_node.add_next_sibling Nokogiri::XML::Node.new "w:br", display_node.document
          else
            text_part = display_node.dup
            text_part.content = part
            display_node.add_next_sibling text_part
          end
        end
      end
    end

    # handles direct addition of WordML to the document template
    class WordML < Struct.new(:xml)
      include Sablon::Content
      def self.id; :word_ml end
      def self.wraps?(value) false end

      def initialize(value)
        super Nokogiri::XML.fragment(value)
      end

      def append_to(paragraph, display_node, env)
        # if all nodes are inline then add them to the existing paragraph
        # otherwise replace the paragraph with the new content.
        if all_inline?
          pr_tag = display_node.parent.at_xpath('./w:rPr')
          add_siblings_to(display_node.parent, pr_tag)
          display_node.parent.remove
        else
          add_siblings_to(paragraph)
          paragraph.remove
        end
      end

      # This allows proper equality checks with other WordML content objects.
      # Due to the fact the `xml` attribute is a live Nokogiri object
      # the default `==` comparison returns false unless it is the exact
      # same object being compared. This method instead checks if the XML
      # being added to the document is the same when the `other` object is
      # an instance of the WordML content class.
      def ==(other)
        if other.class == self.class
          xml.to_s == other.xml.to_s
        else
          super
        end
      end

      private

      # Returns `true` if all of the xml nodes to be inserted are
      def all_inline?
        (xml.children.map(&:node_name) - inline_tags).empty?
      end

      # Array of tags allowed to be a child of the w:p XML tag as defined
      # by the Open XML specification
      def inline_tags
        %w[w:bdo w:bookmarkEnd w:bookmarkStart w:commentRangeEnd
           w:commentRangeStart w:customXml
           w:customXmlDelRangeEnd w:customXmlDelRangeStart
           w:customXmlInsRangeEnd w:customXmlInsRangeStart
           w:customXmlMoveFromRangeEnd w:customXmlMoveFromRangeStart
           w:customXmlMoveToRangeEnd w:customXmlMoveToRangeStart
           w:del w:dir w:fldSimple w:hyperlink w:ins w:moveFrom
           w:moveFromRangeEnd w:moveFromRangeStart w:moveTo
           w:moveToRangeEnd w:moveToRangeStart m:oMath m:oMathPara
           w:pPr w:proofErr w:r w:sdt w:smartTag]
      end

      # Adds the XML to be inserted in the document as siblings to the
      # node passed in. Run properties are merged here because of namespace
      # issues when working with a document fragment
      def add_siblings_to(node, rpr_tag = nil)
        xml.children.reverse.each do |child|
          node.add_next_sibling child
          # merge properties
          next unless rpr_tag
          merge_rpr_tags(child, rpr_tag.children)
        end
      end

      # Merges the provided properties into the run proprties of the
      # node passed in. Properties are only added if they are not already
      # defined on the node itself.
      def merge_rpr_tags(node, props)
        # first assert that all child runs (w:r tags) have a w:rPr tag
        node.xpath('.//w:r').each do |child|
          child.prepend_child '<w:rPr></w:rPr>' unless child.at_xpath('./w:rPr')
        end
        #
        # merge run props, only adding them if they aren't already defined
        node.xpath('.//w:rPr').each do |pr_tag|
          existing = pr_tag.children.map(&:node_name)
          props.map { |pr| pr_tag << pr unless existing.include? pr.node_name }
        end
      end
    end

    # Handles conversion of HTML -> WordML and addition into template
    class HTML < Struct.new(:html_content)
      include Sablon::Content
      def self.id; :html end
      def self.wraps?(value) false end

      def initialize(value)
        super value
      end

      def append_to(paragraph, display_node, env)
        converter = HTMLConverter.new
        word_ml = WordML.new(converter.process(html_content, env))
        word_ml.append_to(paragraph, display_node, env)
      end
    end

    # Handles reading image data and inserting it into the document
    class Image < Struct.new(:name, :data, :properties)
      attr_reader :rid_by_file
      attr_accessor :local_rid

      def self.id; :image end
      def self.wraps?(value) false end

      def inspect
        "#<Image #{name}:#{@rid_by_file}>"
      end

      def initialize(src, attributes = {})
        attributes = Hash[attributes.map { |k, v| [k.to_s, v] }]
        # If the src object is readable, use it as such otherwise open
        # and read the content
        if src.respond_to?(:read)
          name, img_data = process_readable(src, attributes)
        else
          name = File.basename(src)
          img_data = IO.binread(src)
        end
        #
        super name, img_data
        @attributes = attributes
        @properties = @attributes.fetch("properties", {})

        # rId's are separate for each XML file but I want to be able
        # to reuse the actual image file itself.
        @rid_by_file = {}
      end

      def width
        return unless (width_str = @properties[:width])
        convert_to_emu(width_str)
      end

      def height
        return unless (height_str = @properties[:height])
        convert_to_emu(height_str)
      end

      def append_to(paragraph, display_node, env) end

      private

      # Reads the data and attempts to find a filename from either the
      # attributes hash or a #filename method on the src object itself.
      # A filename is required inorder for MS Word to know the content type.
      def process_readable(src, attributes)
        if attributes['filename']
          name = attributes['filename']
        elsif src.respond_to?(:filename)
          name = src.filename
        else
          begin
            name = File.basename(src)
          rescue TypeError
            raise ArgumentError, "Error: Could not determine filename from src, try: `Sablon.content(:image, readable_obj, filename: '...')`"
          end
        end
        #
        [File.basename(name), src.read]
      end

      # Convert centimeters or inches to Word specific emu format
      def convert_to_emu(dim_str)
        value, unit = dim_str.match(/(^\.?\d+\.?\d*)(\w+)/).to_a[1..-1]
        value = value.to_f

        if unit == "cm"
          value = value * 360000
        elsif unit == "in"
          value = value * 914400
        else
          throw ArgumentError, "Unsupported unit '#{unit}', only 'cm' and 'in' are permitted."
        end

        value.round()
      end
    end

    # Handles reading a docx file used like a Ruby on Rails partial
    class Partial
      include Sablon::Content

      def self.id; :partial end
      def self.wraps?(value) false end

      def inspect
        "#<Partial file=#{@filename}>"
      end

      def initialize(src)
        # Read from a file on disk or an IO-like object stored in memory
        @filename = src
        @data = (src.respond_to?(:read) ? src.read : IO.binread(src)).freeze
      end

      def append_to(paragraph, display_node, env)
        # TODO:
        #  * Use DOM to pull in other things used by document.xml such as
        #    images, lists, footnotes, endnotes, links, bookmarks, etc.
        #    Styles will not be ported across.
        #  * Update the existing document.xml with adjusted unique identifiers
        #  * Some complications may arise if the patial is used in the same
        #    document more than once and it has the above ported features.
        local_dom = process_partial(env)
        update_document_relationships(env.document, local_dom)

        # Use WordML to handle the content injection, this is going to be
        # to be a block level replacement 99% of the time since there is
        # usually a paragraph or table at this level and a majority of the
        # content defined as inline isn't allowed to be a child of the
        # body tag.
        body = local_dom.zip_contents['word/document.xml'].xpath('//w:body')
        WordML.new(body.children).append_to(paragraph, display_node, env)
      end

      private

      def process_partial(env)
        # Process the partial using the current context, we only care about
        # the document.xml entry so we don't process anything else
        document = nil
        Zip::File.open_buffer(StringIO.new(@data)) do |z|
          document = Sablon::DOM::Model.new(z)
        end
        document.current_entry = 'word/document.xml'

        xml = document.zip_contents[document.current_entry]
        local_env = Sablon::Environment.new(document, env.context)

        processors = Template.get_processors(document.current_entry)
        processors.each { |processor| processor.process(xml, local_env) }
        #
        document
      end

      def update_document_relationships(env_dom, partial_dom)
        xml = partial_dom.zip_contents['word/document.xml']
        xml_str = xml.to_s

        # Copy over the relevant relationships, if the rel is in the
        # media folder we copy over the new media as well. rId values
        # can appear in multiple elements and places so we use a somewhat
        # general regex to find what nodes to search for
        elms = xml_str.scan(/<(\w+):(\w+) .+"rId\d+".*>/).uniq
        xmlns = elms.flat_map { |n, _| xml_str.scan(/xmlns:(#{n})="(.+?)"/) }
        xmlns = Hash[xmlns]
        nodes = elms.flat_map do |n, e|
          xml.xpath("//#{n}:#{e}", "#{n}": xmlns[n])
        end

        # now that we have the nodes we need to add the references to the
        # current DOM, we don't know what attribute contains the rId so we
        # simply loop oiver it until we find something matching the pattern
        # and then add it to the current document, changing it's value. If
        # the value isn't in the relationship file then we assume it's a
        # false positive.
        nodes.each { |n| duplicate_relationship(env_dom, partial_dom, n) }
      end

      def duplicate_relationship(env_dom, partial_dom, node)
        possible_rids = node.attributes.values.select { |a| a.value =~ /rId\d/ }
        possible_rids.each do |a|
          next unless (rel = partial_dom.find_relationship_by('Id', a.value))
          attrs = Hash[rel.attributes.map { |k, v| [k, v.value] }]
          # copy any media added, ensuring we don't overwrite a file
          if rel['Target'] =~ /media/
            attrs['Target'] = copy_media(env_dom, partial_dom, rel['Target'])
          end
          a.value = env_dom.add_relationship(attrs)
        end
      end

      def copy_media(env_dom, partial_dom, target)
        name = File.basename(target)
        names = env_dom.zip_contents.keys.map { |fn| File.basename(fn) }
        pattern = "^(\\d+)-#{name}"
        val = names.collect { |fn| fn.match(pattern).to_a[1].to_i }.max
        #
        new_name = "media/#{val + 1}-#{name}"
        env_dom.zip_contents["word/#{new_name}"] = partial_dom.zip_contents["word/#{target}"]
        new_name
      end
        end
      end
    end

    register Sablon::Content::String
    register Sablon::Content::WordML
    register Sablon::Content::HTML
    register Sablon::Content::Image
    register Sablon::Content::Partial
  end
end
