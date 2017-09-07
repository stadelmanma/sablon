require "sablon/html/ast_builder"

module Sablon
  class HTMLConverter
    # A top level abstract class to handle common logic for all AST nodes
    class Node
      #attr_reader :children
      PROPERTIES = [].freeze
      FORCE_TRANSFER = [].freeze

      def self.node_name
        @node_name ||= name.split('::').last
      end

      # Returns a hash defined on the configuration object by default. However,
      # this method can be overridden by subclasses to return a different
      # node's style conversion config (i.e. :run) or a hash unrelated to the
      # config itself. The config object is used for all built-in classes to
      # allow for end-user customization via the configuration object
      def self.style_conversion
        # converts camelcase to underscored
        key = node_name.gsub(/([a-z])([A-Z])/, '\1_\2').downcase.to_sym
        Sablon::Configuration.instance.defined_style_conversions.fetch(key, {})
      end

      # maps the CSS style property to it's OpenXML equivalent. Not all CSS
      # properties have an equivalent, nor share the same behavior when
      # defined on different node types (Paragraph, Table and Run).
      def self.process_properties(properties)
        # process the styles as a hash and store values
        style_attrs = {}
        properties.each do |key, value|
          unless key.is_a? Symbol
            key, value = *convert_style_property(key.strip, value.strip)
          end
          style_attrs[key] = value if key
        end
        style_attrs
      end

      # handles conversion of a single attribute allowing recursion through
      # super classes. If the key exists and conversion is succesful a
      # symbol is returned to avoid conflicts with a CSS prop sharing the
      # same name. Keys without a conversion class are returned as is
      def self.convert_style_property(key, value)
        if style_conversion.key?(key)
          key, value = style_conversion[key].call(value)
          key = key.to_sym if key
          [key, value]
        elsif self == Node
          [key, value]
        else
          superclass.convert_style_property(key, value)
        end
      end

      def accept(visitor)
        visitor.visit(self)
      end

      # Simplifies usage at call sites
      def transferred_properties
        @properties.transferred_properties
      end
    end

    class NodeProperties
      attr_reader :transferred_properties
      attr_accessor :force_transfer

      def self.paragraph(properties)
        obj = new('w:pPr', properties, Paragraph::PROPERTIES)
        obj.force_transfer = Paragraph::FORCE_TRANSFER
        #
        obj
      end

      def self.run(properties)
        obj = new('w:rPr', properties, Run::PROPERTIES)
        obj.force_transfer = Run::FORCE_TRANSFER
        #
        obj
      end

      def self.table(properties)
        obj = new('w:tblPr', properties, Table::PROPERTIES)
        obj.force_transfer = Table::FORCE_TRANSFER
        #
        obj
      end

      def self.table_row(properties)
        obj = new('w:trPr', properties, TableRow::PROPERTIES)
        obj.force_transfer = TableRow::FORCE_TRANSFER
        #
        obj
      end

      def self.table_cell(properties)
        obj = new('w:tcPr', properties, TableCell::PROPERTIES)
        obj.force_transfer = TableCell::FORCE_TRANSFER
        #
        obj
      end

      def initialize(tagname, properties, whitelist)
        @tagname = tagname
        filter_properties(properties, whitelist)
      end

      def inspect
        @properties.map { |k, v| v ? "#{k}=#{v}" : k }.join(';')
      end

      def [](key)
        @properties[key]
      end

      def []=(key, value)
        @properties[key] = value
      end

      def to_docx
        "<#{@tagname}>#{process}</#{@tagname}>" unless @properties.empty?
      end

      private

      # processes properties adding those on the whitelist to the
      # properties instance variable and those not to the transferred_properties
      # isntance variable
      def filter_properties(properties, whitelist)
        @transferred_properties = {}
        @properties = {}
        #
        properties.each do |key, value|
          if whitelist.include? key.to_s
            @properties[key] = value
            @transferred_properties[key] = value if force_transfer.include? key.to_s
          else
            @transferred_properties[key] = value
          end
        end
      end

      # processes attributes defined on the node into wordML property syntax
      def process
        @properties.map { |k, v| transform_attr(k, v) }.join
      end

      # properties that have a list as the value get nested in tags and
      # each entry in the list is transformed. When a value is a hash the
      # keys in the hash are used to explicitly build the XML tag attributes.
      def transform_attr(key, value)
        if value.is_a? Array
          sub_attrs = value.map do |sub_prop|
            sub_prop.map { |k, v| transform_attr(k, v) }
          end
          "<w:#{key}>#{sub_attrs.join}</w:#{key}>"
        elsif value.is_a? Hash
          props = value.map { |k, v| format('w:%s="%s"', k, v) if v }
          "<w:#{key} #{props.compact.join(' ')} />"
        else
          value = format('w:val="%s" ', value) if value
          "<w:#{key} #{value}/>"
        end
      end
    end

    class Collection < Node
      attr_reader :nodes
      def initialize(nodes)
        @nodes = nodes
      end

      def accept(visitor)
        super
        @nodes.each do |node|
          node.accept(visitor)
        end
      end

      def to_docx
        nodes.map(&:to_docx).join
      end

      def inspect
        "[#{nodes.map(&:inspect).join(', ')}]"
      end
    end

    class Root < Collection
      def initialize(env, node)
        # strip text nodes from the root level element, these are typically
        # extra whitespace from indenting the markup
        node.search('./text()').remove

        # convert children from HTML to AST nodes
        super(ASTBuilder.html_to_ast(env, node.children, {}))
      end

      def grep(pattern)
        visitor = GrepVisitor.new(pattern)
        accept(visitor)
        visitor.result
      end

      def inspect
        "<Root: #{super}>"
      end
    end

    class Paragraph < Node
      PROPERTIES = %w[framePr ind jc keepLines keepNext numPr
                      outlineLvl pBdr pStyle rPr sectPr shd spacing
                      tabs textAlignment].freeze
      attr_accessor :runs

      def initialize(env, node, properties)
        properties = self.class.process_properties(properties)
        @properties = NodeProperties.paragraph(properties)
        #
        trans_props = transferred_properties
        @runs = ASTBuilder.html_to_ast(env, node.children, trans_props)
        @runs = Collection.new(@runs)
      end

      def to_docx
        "<w:p>#{@properties.to_docx}#{runs.to_docx}</w:p>"
      end

      def accept(visitor)
        super
        runs.accept(visitor)
      end

      def inspect
        "<Paragraph{#{@properties[:pStyle]}}: #{runs.inspect}>"
      end
    end

    class Caption < Paragraph
      def initialize(env, node, properties)
        super
        type = node['type'].capitalize
        pseudo_nodes = <<-html.strip
          #{type}
          <ins placeholder=' #'>SEQ #{type} # \" #\"</ins>
        html
        pseudo_nodes = Nokogiri::XML.fragment(pseudo_nodes)
        #
        bookmark = Bookmark.new(env, pseudo_nodes, properties)
        @runs.nodes.insert(0, bookmark)
      end
    end

    class Table < Node
      PROPERTIES = %w[jc shd tblBorders tblCaption tblCellMar tblCellSpacing
                      tblInd tblLayout tblLook tblOverlap tblpPr tblStyle
                      tblStyleColBandSize tblStyleRowBandSize tblW].freeze
      FORCE_TRANSFER = %w[shd].freeze

      def initialize(env, node, properties)
        properties = self.class.process_properties(properties)
        @properties = NodeProperties.table(properties)
        #
        trans_props = transferred_properties
        @nodes = ASTBuilder.html_to_ast(env, node.children, trans_props)
        @nodes = Collection.new(@nodes)
      end

      def to_docx
        "<w:tbl>#{@properties.to_docx}#{@nodes.to_docx}</w:tbl>"
      end

      def accept(visitor)
        super
        @nodes.accept(visitor)
      end

      def inspect
        "<#{self.class.node_name}{#{@properties.inspect}}: #{@nodes.inspect}>"
      end
    end

    class TableRow < Table
      PROPERTIES = %w[cantSplit hidden jc tblCellSpacing tblHeader
                      trHeight tblPrEx].freeze

      def initialize(env, node, properties)
        properties = self.class.process_properties(properties)
        @properties = NodeProperties.table_row(properties)
        #
        trans_props = transferred_properties
        @nodes = ASTBuilder.html_to_ast(env, node.children, trans_props)
        @nodes = Collection.new(@nodes)
      end

      def to_docx
        "<w:tr>#{@properties.to_docx}#{@nodes.to_docx}</w:tr>"
      end
    end

    class TableCell < Table
      PROPERTIES = %w[gridSpan hideMark noWrap shd tcBorders tcFitText
                      tcMar tcW vAlign vMerge].freeze
      WORD_ML_TAG = 'w:tc'.freeze

      def initialize(env, node, properties)
        properties = self.class.process_properties(properties)
        @properties = NodeProperties.table_cell(properties)
        #
        trans_props = transferred_properties
        @nodes = Paragraph.new(env, node, trans_props)
      end

      def to_docx
        "<w:tc>#{@properties.to_docx}#{@nodes.to_docx}</w:tc>"
      end
    end

    # Manages the child nodes of a list type tag
    class List < Collection
      def initialize(env, node, properties)
        # intialize values
        @list_tag = node.name
        #
        if node.ancestors(".//#{@list_tag}").length.zero?
          # Only register a definition when upon the first list tag encountered
          @definition = env.numbering.register(properties[:pStyle])
        end

        # update attributes of all child nodes
        transfer_node_attributes(node.children, node.attributes)

        # Move any list tags that are a child of a list item up one level
        process_child_nodes(node)

        # strip text nodes from the list level element, this is typically
        # extra whitespace from indenting the markup
        node.search('./text()').remove

        # convert children from HTML to AST nodes
        super(ASTBuilder.html_to_ast(env, node.children, properties))
      end

      def inspect
        "<List: #{super}>"
      end

      private

      # handles passing all attributes on the parent down to children
      def transfer_node_attributes(nodes, attributes)
        nodes.each do |child|
          # update all attributes
          merge_attributes(child, attributes)

          # set attributes specific to list items
          if @definition
            child['pStyle'] = @definition.style
            child['numId'] = @definition.numid
          end
          child['ilvl'] = child.ancestors(".//#{@list_tag}").length - 1
        end
      end

      # merges parent and child attributes together, preappending the parent's
      # values to allow the child node to override it if the value is already
      # defined on the child node.
      def merge_attributes(child, parent_attributes)
        parent_attributes.each do |name, par_attr|
          child_attr = child[name] ? child[name].split(';') : []
          child[name] = par_attr.value.split(';').concat(child_attr).join('; ')
        end
      end

      # moves any list tags that are a child of a list item tag up one level
      # so they become a sibling instead of a child
      def process_child_nodes(node)
        node.xpath("./li/#{@list_tag}").each do |list|
          # transfer attributes from parent now because the list tag will
          # no longer be a child and won't inheirit them as usual
          transfer_node_attributes(list.children, list.parent.attributes)
          list.parent.add_next_sibling(list)
        end
      end
    end

    # Sets list item specific attributes registered on the node
    class ListParagraph < Paragraph
      def initialize(env, node, properties)
        list_props = {
          pStyle: node['pStyle'],
          numPr: [{ ilvl: node['ilvl'] }, { numId: node['numId'] }]
        }
        properties = properties.merge(list_props)
        super
      end

      private

      def transferred_properties
        super
      end
    end

    class Footnote < Node
      attr_reader :placeholder
      #
      Reference = Struct.new(:ref_id) do
        def accept(*_); end

        def inspect
          "<footnoteRef>"
        end

        def to_docx
          <<-wordml.gsub(/^\s*/, '').delete("\n")
            <w:r>
              <w:rPr>
                <w:rStyle w:val="FootnoteReference"/>
              </w:rPr>
              <w:footnoteRef/>
            </w:r>
          wordml
        end
      end

      def initialize(env, node, properties)
        @placeholder = placeholder
        @paragraph = Paragraph.new(env, node, properties)
        @paragraph.runs.nodes.unshift(Reference.new)
        env.footnotes << self
      end

      def ref_id
        @attributes['w:id']
      end

      def ref_id=(value)
        @attributes = { 'w:id' => value }
      end

      def to_docx
        "<w:footnote>#{@paragraph.to_docx}</w:footnote>"
      end

      def accept(visitor)
        super
        @paragraph.accept(visitor)
      end

      def inspect
        "<Footnote: #{@paragraph.inspect}>"
      end
    end

    class FootnoteReference < Run
      attr_reader :placeholder
      #
      Reference = Struct.new(:ref_id) do
        def accept(*_); end

        def inspect
          "id=#{ref_id}"
        end

        def to_docx
          id_str = "w:id=\"#{ref_id}\"" if ref_id
          "<w:footnoteReference #{id_str}/>"
        end
      end

      def initialize(_env, node, properties)
        @properties = NodeProperties.run(properties)
        @placeholder = node['placeholder']
        @ref = Reference.new(node['id'])
        env.footnotes.new_references << self unless node['id']
      end

      def ref_id=(value)
        @children = Reference.new(value)
      end

      def inspect
        "<FootnoteReference{#{@properties.inspect}}: #{@ref.inspect}>"
      end

      def to_docx
        "<w:r>#{@properties.to_docx}#{@ref.to_docx}</w:r>"
      end
    end

    class Bookmark < Collection
      attr_reader :name
      BookmarkTag = Struct.new(:type, :id, :name) do
        def accept(*_); end

        def inspect
          "<Bookmark#{type.capitalize}{id=#{id};name=#{name}}>"
        end

        def to_docx
          "<w:bookmark#{type.capitalize}#{attr_str}/>"
        end

        def attr_str
          attrs = { 'w:id' => id, 'w:name' => name }
          attrs = attrs.map { |k, v| format('%s="%s"', k, v) if v }
          " #{attrs.compact.join(' ')}"
        end
      end

      def initialize(env, node, _properties)
        @name = node['name']
        @nodes[0] = BookmarkTag.new('start', value, @name)
        @nodes.concat ASTBuilder.html_to_ast(env, node.children, {})
        @nodes << BookmarkTag.new('end', nil, nil)
        env.bookmarks << self
      end

      def id=(value)
        @nodes[0] = BookmarkTag.new('start', value, @name)
        @nodes[-1] = BookmarkTag.new('end', value, nil)
      end
    end

    class ComplexField < Collection
      def initialize(_env, node, properties)
        @nodes = [
          FldChar.new(properties, 'begin'),
          InstrText.new(properties, node.text),
          FldChar.new(properties, 'separate'),
          Run.new(properties, node['placeholder'] || ''),
          FldChar.new(properties, 'end')
        ]
      end
    end

    # Create a run of text in the document
    class Run < Node
      PROPERTIES = %w[b i caps color dstrike emboss imprint highlight outline
                      rStyle shadow shd smallCaps strike sz u vanish
                      vertAlign].freeze
      attr_reader :string

      def initialize(_env, node, properties)
        properties = self.class.process_properties(properties)
        @properties = NodeProperties.run(properties)
        @string = node.text
      end

      def to_docx
        "<w:r>#{@properties.to_docx}#{text}</w:r>"
      end

      def inspect
        "<Run{#{@properties.inspect}}: #{string}>"
      end

      private

      def text
        content = @string.tr("\u00A0", ' ')
        "<w:t xml:space=\"preserve\">#{content}</w:t>"
      end
    end

    # isn't meant to be created directly from an HTML node
    class Fldchar < Run
      #
      CharType = Struct.new(:type) do
        def accept(*_); end

        def inspect
          type
        end

        def to_docx
          "<w:fldChar w:fldCharType=\"#{type}\"/>"
        end
      end

      def initialize(properties, type)
        @type = CharType.new(type)
        @properties = NodeProperties.run(properties)
      end

      def inspect
        "<Fldchar{#{@properties.inspect}}: #{@type.inspect}>"
      end

      def to_docx
        "<w:r>#{@properties.to_docx}#{@type.to_docx}</w:r>"
      end
    end

    # isn't meant to be created directly from an HTML node
    class InstrText < Run
      #
      Instructions = Struct.new(:content) do
        def accept(*_); end

        def inspect
          content
        end

        def to_docx
          "<w:instrText xml:space=\"preserve\"> #{content} </w:instrText>"
        end
      end

      def initialize(properties, content)
        @properties = NodeProperties.run(properties)
        @instr = Instructions.new(content)
      end

      def inspect
        "<InstrText{#{@properties.inspect}}: #{@instr.inspect}>"
      end

      def to_docx
        "<w:r>#{@properties.to_docx}#{@instr.to_docx}</w:r>"
      end
    end

    # Creates a blank line in the word document
    class Newline < Node
      def initialize(*); end

      def to_docx
        "<w:r><w:br/></w:r>"
      end

      def inspect
        "<Newline>"
      end
    end
  end
end
