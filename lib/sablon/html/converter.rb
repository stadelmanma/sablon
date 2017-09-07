require "sablon/html/ast"
require "sablon/html/visitor"

module Sablon
  class HTMLConverter
    def process(input, env)
      @env = env
      ast = processed_ast(input)
      # update references before hard conversion into docx string
      @env.footnotes.update_refereces
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
