module Sablon
  class HTMLConverter
    class Node
      PROPERTIES = [].freeze
      def accept(visitor)
        visitor.visit(self)
      end

      def self.node_name
        @node_name ||= name.split('::').last
      end

      private

      def filter_properties(properties)
        props = properties.map do |key, value|
          next unless self.class::PROPERTIES.include? key
          [key, value]
        end
        # filter out nils and return hash
        Hash[props.reject { |pair| pair.nil? }]
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
      attr_accessor :runs
      def initialize(properties, runs)
        @properties = filter_properties(properties)
        @runs = runs
      end

      def to_docx
        "<w:p>#{ppr_docx}#{runs.to_docx}</w:p>"
      end

      def accept(visitor)
        super
        runs.accept(visitor)
      end

      def inspect
        "<Paragraph{#{@properties['pStyle']}}: #{runs.inspect}>"
      end

      private

      def ppr_docx
        "<w:pPr>#{process_properties}</w:pPr>" unless @properties.empty?
      end
    end

    class Run < Node
      PROPERTIES = %w[b i caps color dstrike emboss imprint highlight outline
                      rStyle shadow shd smallCaps strike sz u vanish
                      vertAlign].freeze
      attr_reader :string
      def initialize(string, properties)
        @properties = filter_properties(properties)
        @string = string
      end

      def to_docx
        "<w:r>#{rpr_docx}#{text}</w:r>"
      end

      def inspect
        rpr_str = @properties.map { |k, v| v ? "#{k}=#{v}" : k }.join(';')
        "<Run{#{rpr_str}}: #{string}>"
      end

      private

      def rpr_docx
        "<w:rPr>#{process_properties}</w:rPr>" unless @properties.empty?
      end

      def text
        content = @string.tr("\u00A0", ' ')
        "<w:t xml:space=\"preserve\">#{content}</w:t>"
      end
    end

    class Newline < Node
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
      attr_accessor :rows

      def initialize(properties, rows)
        @properties = filter_properties(properties)
        @rows = rows
      end

      def to_docx
        "<w:tbl>#{tblpr_docx}#{rows.to_docx}</w:tbl>"
      end

      def accept(visitor)
        super
        rows.accept(visitor)
      end

      def inspect
        tblpr_str = @properties.map { |k, v| v ? "#{k}=#{v}" : k }.join(';')
        "<Table{#{tblpr_str}}: #{rows.inspect}>"
      end

      private

      def tblpr_docx
        "<w:tblPr>#{process_properties}</w:tblPr>" unless @properties.empty?
      end
    end

    class TableRow < Node
      PROPERTIES = %w[cantSplit hidden jc tblCellSpacing tblHeader
                      trHeight tblPrEx].freeze
      attr_accessor :cells

      def initialize(properties, cells)
        @properties = filter_properties(properties)
        @cells = cells
      end

      def to_docx
        "<w:tr>#{trpr_docx}#{cells.to_docx}</w:tr>"
      end

      def accept(visitor)
        super
        cells.accept(visitor)
      end

      def inspect
        trpr_str = @properties.map { |k, v| v ? "#{k}=#{v}" : k }.join(';')
        "<TableRow{#{trpr_str}}: #{cells.inspect}>"
      end

      private

      def trpr_docx
        "<w:trPr>#{process_properties}</w:trPr>" unless @properties.empty?
      end
    end

    class TableCell < Node
      PROPERTIES = %w[gridSpan hideMark noWrap shd tcBorders tcFitText
                      tcMar tcW vAlign vMerge].freeze
      attr_accessor :paragraph

      def initialize(properties, runs)
        @properties = filter_properties(properties)
        @paragraph = Paragraph.new(properties, runs)
      end

      def to_docx
        "<w:tc>#{tcpr_docx}#{paragraph.to_docx}</w:tc>"
      end

      def accept(visitor)
        super
        paragraph.accept(visitor)
      end

      def inspect
        tcpr_str = @properties.map { |k, v| v ? "#{k}=#{v}" : k }.join(';')
        "<TableCell{#{tcpr_str}}: #{paragraph.inspect}>"
      end

      private

      def tcpr_docx
        "<w:tcPr>#{process_properties}</w:tcPr>" unless @properties.empty?
      end
    end
  end
end
