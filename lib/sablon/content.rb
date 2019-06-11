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
        # Create a local DOM for the partial and then update various
        # aspects the main document to work with the injected content
        local_dom = process_partial(env)
        update_document_relationships(env.document, local_dom)
        update_document_content_types(env.document, local_dom)
        update_list_definitions(env.document, local_dom)
        update_bookmarks(env.document, local_dom)
        update_footnotes(env.document, local_dom)
        update_endnotes(env.document, local_dom)

        # Use WordML to handle the content injection, this is going to be
        # to be a block level replacement 99% of the time since there is
        # usually a paragraph or table at this level and a majority of the
        # content defined as inline isn't allowed to be a child of the
        # body tag.
        WordML.new(children(local_dom)).append_to(paragraph, display_node, env)
      end

      private

      # Array of tags removed from the body XML prior to injection into
      # the document, the #name method does not include the namespace prefix
      # so we drop it here.
      def excluded_tags
        %w[sectPr]
      end

      # extract valid child tags from the partial's body XML
      def children(local_dom)
        body = local_dom.zip_contents['word/document.xml'].xpath('//w:body')
        nodes = body.children.reject { |n| excluded_tags.include? n.name }
        Nokogiri::XML::NodeSet.new(body.document, nodes)
      end

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
        val = names.collect { |n| n.match(/^(\d+)-#{name}/).to_a[1].to_i }.max
        #
        new_name = "media/#{val + 1}-#{name}"
        env_dom.zip_contents["word/#{new_name}"] = partial_dom.zip_contents["word/#{target}"]
        new_name
      end

      def update_document_content_types(env_dom, partial_dom)
        # update any Extension definitions that are missing from the
        # parent document
        xml = partial_dom.zip_contents['[Content_Types].xml']
        xml.css('Default[Extension]').each do |ctype|
          env_dom.add_content_type(ctype['Extension'], ctype['ContentType'])
        end
      end

      def update_list_definitions(env_dom, partial_dom)
        # determine max numid and max abstract ID present in parent document
        main_xml = env_dom.zip_contents['word/numbering.xml']
        numb = env_dom['word/numbering.xml']
        max_numid = numb.max_attribute_value('//w:num', 'w:numId')
        max_defid = numb.max_attribute_value('//w:abstractNum',
                                             'w:abstractNumId')

        # Collect all the numId elements in the partials document.xml
        # that will be updated with corrected IDs later on
        xml = partial_dom.zip_contents['word/document.xml']
        numid_nodes = xml.xpath('//w:numId')

        # Copy all list definitions in use within the partial into the
        # parent document
        xml = partial_dom.zip_contents['word/numbering.xml']
        num_id_mapping = {}
        numid_nodes.map { |n| n['w:val'] }.uniq.each do |numid|
          num_node = xml.at_xpath("//w:num[@w:numId='#{numid}']")
          id_node = num_node.at_xpath('./w:abstractNumId')
          def_node = xml.at_xpath(
            "//w:abstractNum[@w:abstractNumId='#{id_node['w:val']}']"
          )

          # remove the unique idenifier since I don't want to try and
          # regenerate a valid value
          def_node.xpath('./w:nsid').each(&:remove)

          # adjust attribute values and update mapping
          num_node['w:numId'] = (max_numid += 1)
          id_node['w:val'] = (max_defid += 1)
          def_node['w:abstractNumId'] = id_node['w:val']
          num_id_mapping[numid] = max_numid

          # Copy definitions from partial into parent
          node = main_xml.xpath('//w:abstractNum').last
          node.add_next_sibling(def_node)
          node = main_xml.xpath('//w:num').last
          node.add_next_sibling(num_node)
        end

        # update w:numId nodes inside the document.xml for the partial
        # according to a mapping
        numid_nodes.each { |n| n['w:val'] = num_id_mapping[n['w:val']] }
      end

      def update_bookmarks(env_dom, partial_dom)
        # determine the maximum boookmark ID value and get a list of all
        # names in use
        xml = env_dom.zip_contents['word/document.xml']
        doc = env_dom['word/document.xml']
        max_id = doc.max_attribute_value(xml, '//w:bookmarkStart', 'w:id')
        used_names = xml.xpath('//w:bookmarkStart').map { |n| n['w:name'] }

        # locate all bookmark start and end tags
        xml = partial_dom.zip_contents['word/document.xml']
        st_nodes = xml.xpath('//w:bookmarkStart')
        en_nodes = xml.xpath('//w:bookmarkEnd')

        # Collect a unique list of names and IDs to update via a mapping
        ids = st_nodes.map { |n| n['w:id'] }
        names = st_nodes.map { |n| n['w:name'] }
        id_mapping = Hash[ids.map { |v| [v, (max_id += 1)] }]
        name_mapping = Hash[names.map { |v| [v, v] }]

        # check for any names that exist in both documents and adjust
        # mapping value
        (names & used_names).each do |name|
          pat = /^#{name}_(\d+)/
          val = used_names.collect { |n| n.match(pat).to_a[1].to_i }.max
          name_mapping[name] = "#{name}_#{val + 1}"
        end

        # update nodes
        (st_nodes + en_nodes).each { |n| n['w:id'] = id_mapping[n['w:id']] }
        st_nodes.each { |n| n['w:name'] = name_mapping[n['w:name']] }

        # update bookmark references, I think these are the only nodes
        # to use them, if the reference can't be found in the mapping then
        # we will leave it alone.
        xml.xpath('//w:instrText').each do |n|
          next unless (mat = n.text.match(/\s*REF (\w+)/))
          next unless (bk_name = name_mapping[mat[1]])
          n.content = n.text.sub(/\s*REF (\w+)/, " REF #{bk_name}")
        end
      end

      def update_footnotes(env_dom, partial_dom)
        update_notes(env_dom, partial_dom, 'footnote')
      end

      def update_endnotes(env_dom, partial_dom)
        update_notes(env_dom, partial_dom, 'endnote')
      end

      def update_notes(env_dom, partial_dom, kind)
        # determine the maximum footnote ID value
        xml = env_dom.zip_contents["word/#{kind}s.xml"]
        entry = env_dom["word/#{kind}s.xml"]
        if entry.nil?
          raise ContextError,
                "Partial contains #{kind}s but main document does not."
        end
        max_id = entry.max_attribute_value(xml, "//w:#{kind}", 'w:id')

        # collect all footnote references in use inside the partial
        xml = partial_dom.zip_contents['word/document.xml']
        refs = xml.xpath("//w:#{kind}Reference")

        # create an id mapping and then start updating nodes and copying
        # the footnote definitions
        main_xml = env_dom.zip_contents["word/#{kind}s.xml"]
        xml = partial_dom.zip_contents["word/#{kind}s.xml"]
        id_mapping = Hash[refs.map { |n| [n['w:id'], (max_id += 1)] }]
        refs.each do |ref|
          node = xml.at_xpath("//w:#{kind}[@w:id='#{ref['w:id']}']")
          ref['w:id'] = id_mapping[ref['w:id']]
          node['w:id'] = id_mapping[node['w:id']]
          #
          last_node = main_xml.xpath("//w:#{kind}").last
          last_node.add_next_sibling(node)
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
