# -*- coding: utf-8 -*-
require "test_helper"

class HTMLConverterASTTest < Sablon::TestCase
  def setup
    super
    env = Sablon::Environment.new(nil)
    @bookmarks = env.bookmarks
    @footnotes = env.footnotes
    @footnotes.instance_variable_set(:@counter, 1)
    #
    @converter = Sablon::HTMLConverter.new
    @converter.instance_variable_set(:@env, env)
  end

  def test_div
    input = '<div>Lorem ipsum dolor sit amet</div>'
    ast = @converter.processed_ast(input)
    assert_equal '<Root: [<Paragraph{pStyle=Normal}: [<Run{}: Lorem ipsum dolor sit amet>]>]>', ast.inspect
  end

  def test_p
    input = '<p>Lorem ipsum dolor sit amet</p>'
    ast = @converter.processed_ast(input)
    assert_equal '<Root: [<Paragraph{pStyle=Paragraph}: [<Run{}: Lorem ipsum dolor sit amet>]>]>', ast.inspect
  end

  def test_b
    input = '<p>Lorem <b>ipsum dolor sit amet</b></p>'
    ast = @converter.processed_ast(input)
    assert_equal '<Root: [<Paragraph{pStyle=Paragraph}: [<Run{}: Lorem >, <Run{b}: ipsum dolor sit amet>]>]>', ast.inspect
  end

  def test_i
    input = '<p>Lorem <i>ipsum dolor sit amet</i></p>'
    ast = @converter.processed_ast(input)
    assert_equal '<Root: [<Paragraph{pStyle=Paragraph}: [<Run{}: Lorem >, <Run{i}: ipsum dolor sit amet>]>]>', ast.inspect
  end

  def test_br_in_strong
    input = '<div><strong>Lorem<br />ipsum<br />dolor</strong></div>'
    par = @converter.processed_ast(input).grep(Sablon::HTMLConverter::Paragraph).first
    assert_equal "[<Run{b}: Lorem>, <Newline>, <Run{b}: ipsum>, <Newline>, <Run{b}: dolor>]", par.children.inspect
  end

  def test_br_in_em
    input = '<div><em>Lorem<br />ipsum<br />dolor</em></div>'
    par = @converter.processed_ast(input).grep(Sablon::HTMLConverter::Paragraph).first
    assert_equal "[<Run{i}: Lorem>, <Newline>, <Run{i}: ipsum>, <Newline>, <Run{i}: dolor>]", par.children.inspect
  end

  def test_nested_strong_and_em
    input = '<div><strong>Lorem <em>ipsum</em> dolor</strong></div>'
    par = @converter.processed_ast(input).grep(Sablon::HTMLConverter::Paragraph).first
    assert_equal "[<Run{b}: Lorem >, <Run{b;i}: ipsum>, <Run{b}:  dolor>]", par.children.inspect
  end

  def test_ignore_last_br_in_div
    input = '<div>Lorem ipsum dolor sit amet<br /></div>'
    par = @converter.processed_ast(input).grep(Sablon::HTMLConverter::Paragraph).first
    assert_equal "[<Run{}: Lorem ipsum dolor sit amet>]", par.children.inspect
  end

  def test_ignore_br_in_blank_div
    input = '<div><br /></div>'
    par = @converter.processed_ast(input).grep(Sablon::HTMLConverter::Paragraph).first
    assert_equal "[]", par.children.inspect
  end

  def test_headings
    input = '<h1>First</h1><h2>Second</h2><h3>Third</h3>'
    ast = @converter.processed_ast(input)
    assert_equal "<Root: [<Paragraph{pStyle=Heading1}: [<Run{}: First>]>, <Paragraph{pStyle=Heading2}: [<Run{}: Second>]>, <Paragraph{pStyle=Heading3}: [<Run{}: Third>]>]>", ast.inspect
  end

  def test_h_with_formatting
    input = '<h1><strong>Lorem</strong> ipsum dolor <em>sit <u>amet</u></em></h1>'
    ast = @converter.processed_ast(input)
    assert_equal "<Root: [<Paragraph{pStyle=Heading1}: [<Run{b}: Lorem>, <Run{}:  ipsum dolor >, <Run{i}: sit >, <Run{i;u=single}: amet>]>]>", ast.inspect
  end

  def test_ul
    input = '<ul><li>Lorem</li><li>ipsum</li></ul>'
    ast = @converter.processed_ast(input)
    assert_equal "<Root: [<List: [<Paragraph{pStyle=ListBullet;numPr=[{:ilvl=>\"0\"}, {:numId=>\"1001\"}]}: [<Run{}: Lorem>]>, <Paragraph{pStyle=ListBullet;numPr=[{:ilvl=>\"0\"}, {:numId=>\"1001\"}]}: [<Run{}: ipsum>]>]>]>", ast.inspect
  end

  def test_ol
    input = '<ol><li>Lorem</li><li>ipsum</li></ol>'
    ast = @converter.processed_ast(input)
    assert_equal "<Root: [<List: [<Paragraph{pStyle=ListNumber;numPr=[{:ilvl=>\"0\"}, {:numId=>\"1001\"}]}: [<Run{}: Lorem>]>, <Paragraph{pStyle=ListNumber;numPr=[{:ilvl=>\"0\"}, {:numId=>\"1001\"}]}: [<Run{}: ipsum>]>]>]>", ast.inspect
  end

  def test_num_id
    ast = @converter.processed_ast('<ol><li>Some</li><li>Lorem</li></ol><ul><li>ipsum</li></ul><ol><li>dolor</li><li>sit</li></ol>')
    assert_equal %w[1001 1001 1002 1003 1003], get_numpr_prop_from_ast(ast, :numId)
  end

  def test_nested_lists_have_the_same_numid
    ast = @converter.processed_ast('<ul><li>Lorem<ul><li>ipsum<ul><li>dolor</li></ul></li></ul></li></ul>')
    assert_equal %w[1001 1001 1001], get_numpr_prop_from_ast(ast, :numId)
  end

  def test_keep_nested_list_order
    input = '<ul><li>1<ul><li>1.1<ul><li>1.1.1</li></ul></li><li>1.2</li></ul></li><li>2<ul><li>1.3<ul><li>1.3.1</li></ul></li></ul></li></ul>'
    ast = @converter.processed_ast(input)
    assert_equal %w[1001], get_numpr_prop_from_ast(ast, :numId).uniq
    assert_equal %w[0 1 2 1 0 1 2], get_numpr_prop_from_ast(ast, :ilvl)
  end

  def test_footnoteref
    input = '<p>Lorem<footnoteref id="2"/></p>'
    ast = @converter.processed_ast(input)
    assert_equal "<Root: [<Paragraph{pStyle=Paragraph}: [<Run{}: Lorem>, <FootnoteReference{rStyle=FootnoteReference}: id=2>]>]>", ast.inspect
  end

  def test_footnote
    input = '<footnote placeholder="test">Lorem Ipsum</footnote>'
    ast = @converter.processed_ast(input)
    #
    assert_equal 1, @footnotes.new_footnotes.length
    assert_equal "2", @footnotes.new_footnotes[0].ref_id
    assert_equal "<Root: [<Footnote{}: <Paragraph{pStyle=FootnoteText}: [<footnoteRef>, <Run{}: Lorem Ipsum>]>>]>", ast.inspect
  end

  def test_bookmark
    input = '<p><bookmark name="test">Lorem</bookmark> Ipsum</p>'
    ast = @converter.processed_ast(input)
    #
    assert_equal "<Root: [<Paragraph{pStyle=Paragraph}: [[<BookmarkStart{id=1;name=test}>, <Run{}: Lorem>, <BookmarkEnd{id=1;name=}>], <Run{}:  Ipsum>]>]>", ast.inspect
    assert_equal @bookmarks.instance_variable_get(:@counter), 1
    assert_equal @bookmarks.instance_variable_get(:@names), ['test']
  end

  private

  # returns the numid attribute from paragraphs
  def get_numpr_prop_from_ast(ast, key)
    values = []
    ast.grep(Sablon::HTMLConverter::ListParagraph).each do |para|
      numpr = para.instance_variable_get('@properties')[:numPr]
      numpr.each { |val| values.push(val[key]) if val[key] }
    end
    values
  end
end
