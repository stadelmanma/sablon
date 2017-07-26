require "sablon/html/ast"
require "sablon/html/visitor"

module Sablon
  class HTMLConverter
    class ASTBuilder
      Layer = Struct.new(:items, :ilvl)

      def initialize(nodes)
        @layers = [Layer.new(nodes, false)]
        @root = Root.new([])
      end

      def to_ast
        @root
      end

      def new_layer(ilvl: false)
        @layers.push Layer.new([], ilvl)
      end

      def next
        current_layer.items.shift
      end

      def push(node)
        @layers.last.items.push node
      end

      def push_all(nodes)
        nodes.each(&method(:push))
      end

      def done?
        !current_layer.items.any?
      end

      def nested?
        ilvl > 0
      end

      def ilvl
        @layers.select { |layer| layer.ilvl }.size - 1
      end

      def emit(node)
        @root.nodes << node
      end

      private

      def current_layer
        if @layers.any?
          last_layer = @layers.last
          if last_layer.items.any?
            last_layer
          else
            @layers.pop
            current_layer
          end
        else
          Layer.new([], false)
        end
      end
    end

    def process(input, env)
      @bookmarks = env.bookmarks
      @numbering = env.numbering
      @footnotes = env.footnotes
      ast = processed_ast(input)
      # update references before hard conversion into docx string
      @footnotes.update_refereces
      ast.to_docx
    end

    def processed_ast(input)
      ast = build_ast(input)
      ast.accept LastNewlineRemoverVisitor.new
      ast
    end

    def build_ast(input)
      doc = Nokogiri::HTML.fragment(input)
      @builder = ASTBuilder.new(doc.children)

      while !@builder.done?
        ast_next_paragraph
      end
      @builder.to_ast
    end

    private

    def initialize
      @numbering = nil
      @new_footnotes = nil
    end

    # Adds the appropriate style class to the node
    def prepare_paragraph(node, properties = {})
      # determine conversion class for table separately.
      node_cls = { 'table' => Table, 'tr' => TableRow, 'td' => TableCell,
                   'th' => TableCell, 'footnote' => Footnote,
                   'caption' => Caption }
      node_cls.default = Paragraph
      # set default styles based on HTML element allowing for h1, h2, etc.
      styles = Hash.new do |hash, key|
        tag, num = key.match(/([a-z]+)(\d*)/)[1..2]
        { 'pStyle' => hash[tag]['pStyle'] + num } if hash.key? tag
      end
      styles.merge!('div' => 'Normal', 'p' => 'Paragraph', 'h' => 'Heading',
                    'table' => nil, 'tr' => nil, 'td' => nil,
                    'footnote' => 'FootnoteText', 'caption' => 'Caption',
                    'ul' => 'ListBullet', 'ol' => 'ListNumber')
      styles['li'] = @definition.style if @definition
      styles.each { |k, v| styles[k] = v ? { 'pStyle' => v } : {} }
      styles['th'] = { 'b' => nil, 'jc' => 'center' }
      #
      unless styles[node.name] || styles.key?(node.name)
        raise ArgumentError, "Don't know how to handle node: #{node.inspect}"
      end
      #
      merge_properties(node, properties, styles[node.name], node_cls[node.name])
    end

    # Adds properties to the run, from the parent, the style node attributes
    # and finally any element specfic properties. A modified properties hash
    # is returned
    def prepare_run(node, properties)
      # HTML element based styles
      styles = {
        'span' => {}, 'br' => {}, 'text' => {},
        'strong' => { 'b' => nil }, 'b' => { 'b' => nil },
        'em' => { 'i' => nil }, 'i' => { 'i' => nil },
        'u' => { 'u' => 'single' },
        's' => { 'strike' => 'true' },
        'sub' => { 'vertAlign' => 'subscript' },
        'sup' => { 'vertAlign' => 'superscript' },
        'ins' => { 'noProof' => nil },
        'footnoteref' => { 'rStyle' => 'FootnoteReference' },
        'bookmark' => {},
        'bgcyan' => { 'highlight' => { val: 'cyan' } },
        'bggreen' => { 'highlight' => { val: 'green' } },
        'bgmagenta' => { 'highlight' => { val: 'magenta' } },
        'bgyellow' => { 'highlight' => { val: 'yellow' } },
        'bgwhite' => { 'highlight' => { val: 'white' } }
      }

      unless styles.key?(node.name)
        raise ArgumentError, "Don't know how to handle node: #{node.inspect}"
      end
      merge_properties(node, properties, styles[node.name], Run)
    end

    def merge_properties(node, par_props, elm_props, ast_class)
      # perform an initial conversion for any leftover CSS props passed in
      properties = par_props.map do |k, v|
        ast_class.convert_style_attr(k, v)
      end
      properties = Hash[properties]

      # Process any styles, defined on the node
      properties.merge!(ast_class.process_style(node['style']))

      # Set the element specific attributes, overriding any other values
      properties.merge(elm_props)
    end

    # handles passing all attributes on the parent down to children
    # preappending parent attributes so child can overwrite if present
    def merge_node_attributes(node, attributes)
      node.children.each do |child|
        attributes.each do |name, atr|
          catr = child[name] ? child[name] : ''
          child[name] = atr.value.split(';').concat(catr.split(';')).join('; ')
        end
      end
    end

    def ast_next_paragraph
      node = @builder.next
      return if node.text?

      properties = prepare_paragraph(node)

      # handle special cases
      if node.name =~ /ul|ol/
        @builder.new_layer ilvl: true
        unless @builder.nested?
          @definition = @numbering.register(properties['pStyle'])
        end
        merge_node_attributes(node, node.attributes)
        @builder.push_all(node.children)
        return
      elsif node.name == 'li'
        properties['numPr'] = [
          { 'ilvl' => @builder.ilvl }, { 'numId' => @definition.numid }
        ]
      elsif node.name == 'caption'
        trans_props = Caption.transferred_properties(properties)
        caption = Caption.new(properties, node['type'], node['name'], ast_runs(node.children, trans_props))
        @bookmarks << caption.bookmark
        @builder.new_layer
        @builder.emit caption
        return
      elsif node.name == 'footnote'
        trans_props = Footnote.transferred_properties(properties)
        @footnotes << Footnote.new(properties, node['placeholder'], ast_runs(node.children, trans_props))
        return
      elsif node.name == 'table'
        @builder.new_layer
        trans_props = Table.transferred_properties(properties)
        @builder.emit Table.new(properties, ast_process_table_rows(node.children, trans_props))
        return
      end

      # create word_ml node
      @builder.new_layer
      trans_props = Paragraph.transferred_properties(properties)
      @builder.emit Paragraph.new(properties, ast_runs(node.children, trans_props))
    end

    def ast_process_table_rows(nodes, properties)
      rows = nodes.map do |node|
        next unless node.name == 'tr' # ignore everything that isn't a row
        local_props = prepare_paragraph(node, properties)
        trans_props = TableRow.transferred_properties(local_props)
        TableRow.new(local_props, ast_process_table_cells(node.children, trans_props))
      end
      Collection.new(rows.compact)
    end

    def ast_process_table_cells(nodes, properties)
      cells = nodes.map do |node|
        # ignore everything that isn't a cell
        next unless node.name == 'td' || node.name == 'th'
        local_props = prepare_paragraph(node, properties)
        para_props = TableCell.transferred_properties(local_props)
        run_props = Paragraph.transferred_properties(para_props)
        paragraph = Paragraph.new(para_props, ast_runs(node.children, run_props))
        TableCell.new(local_props, paragraph)
      end
      Collection.new(cells.compact)
    end

    def ast_runs(nodes, properties)
      runs = nodes.flat_map do |node|
        begin
          local_props = prepare_run(node, properties)
        rescue ArgumentError
          raise unless %w[ul ol p div].include?(node.name)
          merge_node_attributes(node, node.parent.attributes)
          @builder.push(node)
          next nil
        end
        #
        if node.text?
          Run.new(local_props, node.text)
        elsif node.name == 'bookmark'
          child_nodes = ast_runs(node.children, local_props).nodes
          bookmark = Bookmark.new(node['name'], child_nodes)
          @bookmarks << bookmark
          bookmark
        elsif node.name == 'br'
          Newline.new
        elsif node.name == 'footnoteref'
          ref = FootnoteReference.new(local_props, node)
          @footnotes.new_references << ref unless node['id']
          ref
        elsif node.name == 'ins'
          ComplexField.new(local_props, node.text, node['placeholder'])
        else
          ast_runs(node.children, local_props).nodes
        end
      end
      Collection.new(runs.compact)
    end
  end
end
