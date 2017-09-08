# -*- coding: utf-8 -*-
module Sablon
  module Processor
    # Manages footnote usage in the document
    class Footnotes
      attr_reader :new_footnotes
      attr_reader :new_references
      attr_reader :counter
      def self.process(xml_node, env, *args)
        Document.process(xml_node, env, *args)
        # add all new footnotes to the file
        footnote_parent = xml_node.at_xpath('.//w:footnotes')
        env.footnotes.new_footnotes.each do |footnote|
          footnote_parent.add_child(footnote.to_docx(true))
        end
        #
        xml_node
      end

      def initialize_footnotes(xml_node)
        return unless xml_node
        footnotes = xml_node.at_xpath('.//w:footnotes').children
        @counter = footnotes.last.attributes['id'].value.to_i
      end

      # adds a new footnote
      def <<(footnote)
        add_footnote(footnote)
        @placeholder_map[footnote.placeholder] = footnote
      end

      def update_refereces
        @new_references.each do |ref|
          if @placeholder_map[ref.placeholder]
            # remove the footnote from the map as they can only be used once
            footnote = @placeholder_map.delete(ref.placeholder)
            # store used footnote
            @used_footnotes[ref.placeholder] = footnote
          elsif @used_footnotes[ref.placeholder]
            # When a footnote gets reused duplicate the original, this
            # prevent MS Word from thinking the document is corrupt
            footnote = @used_footnotes[ref.placeholder].dup
            add_footnote(footnote)
          else
            raise ContextError, "Bad footnote reference: #{ref.placeholder}"
          end
          ref.ref_id = footnote.ref_id
        end
      end

      private

      def initialize
        @counter = 0
        @new_footnotes = []
        @used_footnotes = {}
        @new_references = []
        @placeholder_map = {}
      end

      # increments counter and stores the new footnote
      def add_footnote(footnote)
        @counter += 1
        footnote.ref_id = @counter.to_s
        @new_footnotes << footnote
      end
    end
  end
end
