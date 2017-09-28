require "sablon/html/ast"
require "sablon/html/visitor"

module Sablon
  class HTMLConverter
    def process(input, env)
      @env = env
      ast = processed_ast(input)
      # update references before hard conversion into docx string, this would
      # be an ideal place to register a "before_to_docx" hook or callback that
      # any processor can tie into to perform behaviour like this. That type
      # of system would also prevent extra logic from being stuffed here where
      # it currently has to be but doesn't belong.
      @env.footnotes.update_refereces
      #
      ast.to_docx
    end

    def processed_ast(input)
      ast = build_ast(input)
      ast.accept LastNewlineRemoverVisitor.new
      ast
    end

    def build_ast(input)
      doc = Nokogiri::HTML.fragment(input)
      Root.new(@env, doc)
    end
  end
end
