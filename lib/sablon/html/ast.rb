module Sablon
  class HTMLConverter
    class Node
      PROPERTIES = [].freeze
      FORCE_TRANSFER = [].freeze
      # styles shared or common logic across all node types go here. Any
      # undefined styles are passed straight through "as is" to the
      # properties hash. Keys that are symbols will not get called directly
      # when processing the style string and are suitable for internal-only
      # usage across different classes.
      STYLE_CONVERSION = {
        'background-color' => lambda { |v|
          return 'shd', { val: 'clear', fill: v.delete('#') }
        },
        border: lambda { |v|
          props = { sz: 2, val: 'single', color: '000000' }
          vals = v.split
          vals[1] = 'single' if vals[1] == 'solid'
          #
          props[:sz] = (2 * Float(vals[0].gsub(/[^\d.]/, '')).ceil).to_s if vals[0]
          props[:val] = vals[1] if vals[1]
          props[:color] = vals[2].delete('#') if vals[2]
          #
          return props
        },
        'text-align' => ->(v) { return 'jc', v }
      }
      STYLE_CONVERSION.default_proc = proc do |hash, key|
        ->(v) { return key, v }
      end
      STYLE_CONVERSION.freeze
      attr_accessor :children

      def self.node_name
        @node_name ||= name.split('::').last
      end

      # maps the CSS style property to it's OpenXML equivalent. Not all CSS
      # properties have an equivalent, nor share the same behavior when
      # defined on different node types (Paragraph, Table and Run).
      def self.process_style(style_str)
        return {} unless style_str
        #
        styles = style_str.split(';').map { |pair| pair.split(':') }
        # process the styles as a hash and store values
        style_attrs = {}
        Hash[styles].each do |key, value|
          key, value = convert_style_attr(key.strip, value.strip)
          style_attrs[key] = value if key
        end
        style_attrs
      end

      # handles conversion of a single attribute allowing recursion through
      # super classes until the node class is reached
      def self.convert_style_attr(key, value)
        if self::STYLE_CONVERSION[key]
          self::STYLE_CONVERSION[key].call(value)
        else
          superclass.convert_style_attr(key, value)
        end
      end

      # creates a hash of all properties that aren't consumed by the node
      # to be passed onto child nodes
      def self.transferred_properties(properties)
        props = properties.map do |key, value|
          next [key, value] if self::FORCE_TRANSFER.include? key
          next if self::PROPERTIES.include? key
          [key, value]
        end
        # filter out nils and return hash
        Hash[props.compact]
      end

      def initialize(properties, children)
        @attributes ||= {}
        @properties = filter_properties(properties)
        @children = children
      end

      def accept(visitor)
        visitor.visit(self)
        children.accept(visitor) if children
      end

      def to_docx
        tag = self.class::WORD_ML_TAG
        attr_str = attributes_to_docx unless @attributes.empty?
        "<#{tag}#{attr_str}>#{properties_to_docx}#{children.to_docx}</#{tag}>"
      end

      def inspect
        pr_str = @properties.map { |k, v| v ? "#{k}=#{v}" : k }.join(';')
        "<#{self.class.node_name}{#{pr_str}}: #{children.inspect}>"
      end

      private

      # removes properties that are not whitelisted by the class
      def filter_properties(properties)
        props = properties.map do |key, value|
          next unless self.class::PROPERTIES.include? key
          [key, value]
        end
        # filter out nils and return hash
        Hash[props.compact]
      end

      # processes attributes defined on the node into wordML property syntax
      def process_properties
        @properties.map { |k, v| transform_attr(k, v) }.join
      end

      # properties that have a list as the value get nested in tags and
      # each entry in the list is transformed. When a value is a hash the
      # keys in the hash are used to explicitly buld the XML tag attributes.
      def transform_attr(key, value)
        if value.is_a? Array
          sub_attrs = value.map do |sub_prop|
            sub_prop.map { |k, v| transform_attr(k, v) }
          end
          "<w:#{key}>#{sub_attrs.join}</w:#{key}>"
        elsif value.is_a? Hash
          props = value.map { |k, v| format('w:%s="%s"', k, v) }
          "<w:#{key} #{props.join(' ')} />"
        else
          value = format('w:val="%s" ', value) if value
          "<w:#{key} #{value}/>"
        end
      end

      def attributes_to_docx
        attrs = @attributes.map { |k, v| format('%s="%s"', k, v) if v }
        " #{attrs.join(' ')}"
      end

      def properties_to_docx
        tag = self.class::WORD_ML_TAG + 'Pr'
        "<#{tag}>#{process_properties}</#{tag}>" unless @properties.empty?
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
      WORD_ML_TAG = 'w:p'.freeze
      STYLE_CONVERSION = {
        'border' => lambda { |v|
          props = Node::STYLE_CONVERSION[:border].call(v)
          #
          return 'pBdr', [
            { top: props }, { bottom: props },
            { left: props }, { right: props }
          ]
        },
        'vertical-align' => ->(v) { return 'textAlignment', v }
      }.freeze
    end

    class Run < Node
      PROPERTIES = %w[b i caps color dstrike emboss imprint highlight outline
                      rStyle shadow shd smallCaps strike sz u vanish
                      vertAlign].freeze
      WORD_ML_TAG = 'w:r'.freeze
      STYLE_CONVERSION = {
        'color' => ->(v) { return 'color', v.delete('#') },
        'font-size' => lambda { |v|
          return 'sz', (2 * Float(v.gsub(/[^\d.]/, '')).ceil).to_s
        },
        'font-style' => lambda { |v|
          return 'b', nil if v =~ /bold/
          return 'i', nil if v =~ /italic/
        },
        'font-weight' => ->(v) { return 'b', nil if v =~ /bold/ },
        'text-decoration' => lambda { |v|
          supported = %w[line-through underline]
          props = v.split
          return props[0], 'true' unless supported.include? props[0]
          return 'strike', 'true' if props[0] == 'line-through'
          return 'u', 'single' if props.length == 1
          return 'u', { val: props[1], color: 'auto' } if props.length == 2
          return 'u', { val: props[1], color: props[2].delete('#') }
        },
        'vertical-align' => lambda { |v|
          return 'vertAlign', 'subscript' if v =~ /sub/
          return 'vertAlign', 'superscript' if v =~ /super/
        }
      }.freeze

      Text = Struct.new(:string) do
        def accept(*_); end

        def inspect
          string
        end

        def to_docx
          content = string.tr("\u00A0", ' ')
          "<w:t xml:space=\"preserve\">#{content}</w:t>"
        end
      end

      def initialize(properties, string)
        super properties, Text.new(string)
      end
    end

    class Footnote < Node
      WORD_ML_TAG = 'w:footnote'.freeze
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

      def initialize(properties, placeholder, children)
        @placeholder = placeholder
        children.nodes.insert(0, Reference.new)
        super({}, Paragraph.new(properties, children))
      end

      def ref_id
        @attributes['w:id']
      end

      def ref_id=(value)
        @attributes = { 'w:id' => value }
      end
    end

    class FootnoteReference < Node
      PROPERTIES = Run::PROPERTIES
      WORD_ML_TAG = Run::WORD_ML_TAG
      STYLE_CONVERSION = Run::STYLE_CONVERSION
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

      def initialize(properties, node)
        @placeholder = node['placeholder']
        super properties, Reference.new(node['id'])
      end

      def ref_id=(value)
        @children = Reference.new(value)
      end
    end

    class Newline < Node
      def initialize; end

      def to_docx
        "<w:r><w:br/></w:r>"
      end

      def inspect
        "<Newline>"
      end
    end

    class Table < Node
      PROPERTIES = %w[jc shd tblBorders tblCaption tblCellMar tblCellSpacing
                      tblInd tblLayout tblLook tblOverlap tblpPr tblStyle
                      tblStyleColBandSize tblStyleRowBandSize tblW].freeze
      FORCE_TRANSFER = %[shd].freeze
      WORD_ML_TAG = 'w:tbl'.freeze
      STYLE_CONVERSION = {
        'border' => lambda { |v|
          props = Node::STYLE_CONVERSION[:border].call(v)
          #
          return 'tblBorders', [
            { top: props }, { start: props }, { bottom: props },
            { end: props }, { insideH: props }, { insideV: props }
          ]
        },
        'margin' => lambda { |v|
          vals = v.split.map { |s| (2 * Float(s.gsub(/[^\d.]/, '')).ceil).to_s }
          #
          props = [vals[0], vals[0], vals[0], vals[0]] if vals.length == 1
          props = [vals[0], vals[1], vals[0], vals[1]] if vals.length == 2
          props = [vals[0], vals[1], vals[2], vals[1]] if vals.length == 3
          props = [vals[0], vals[1], vals[2], vals[3]] if vals.length > 3
          return 'tblCellMar', [
            { top: { w: props[0], type: 'dxa' } },
            { end: { w: props[1], type: 'dxa' } },
            { bottom: { w: props[2], type: 'dxa' } },
            { start: { w: props[3], type: 'dxa' } }
          ]
        },
        'cellspacing' => lambda { |v|
          v = (2 * Float(v.gsub(/[^\d.]/, '')).ceil).to_s
          return 'tblCellSpacing', { w: v, type: 'dxa' }
        },
        'width' => lambda { |v|
          v = (2 * Float(v.gsub(/[^\d.]/, '')).ceil).to_s
          return 'tblW', { w: v, type: 'dxa' }
        }
      }.freeze
    end

    class TableRow < Node
      PROPERTIES = %w[cantSplit hidden jc tblCellSpacing tblHeader
                      trHeight tblPrEx].freeze
      WORD_ML_TAG = 'w:tr'.freeze
    end

    class TableCell < Node
      PROPERTIES = %w[gridSpan hideMark noWrap shd tcBorders tcFitText
                      tcMar tcW vAlign vMerge].freeze
      WORD_ML_TAG = 'w:tc'.freeze
      STYLE_CONVERSION = {
        'border' => lambda { |v|
          value = Table::STYLE_CONVERSION['border'].call(v)[1]
          return 'tcBorders', value
        },
        'colspan' => ->(v) { return 'gridSpan', v },
        'margin' => lambda { |v|
          value = Table::STYLE_CONVERSION['margin'].call(v)[1]
          return 'tcMar', value
        },
        'rowspan' => lambda { |v|
          return 'vMerge', 'restart' if v == 'start'
          return 'vMerge', v if v == 'continue'
          return 'vMerge', nil if v == 'end'
        },
        'vertical-align' => ->(v) { return 'vAlign', v },
        'white-space' => lambda { |v|
          return 'noWrap', nil if v == 'nowrap'
          return 'tcFitText', 'true' if v == 'fit'
        },
        'width' => lambda { |v|
          value = Table::STYLE_CONVERSION['width'].call(v)[1]
          return 'tcW', value
        }
      }.freeze
    end
  end
end
