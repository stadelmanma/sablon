# -*- coding: utf-8 -*-
require "test_helper"
require "support/xml_snippets"


class HTMLConverterTest < Sablon::TestCase
  include XMLSnippets

  def setup
    super
    @env = Sablon::Environment.new(nil)
    @numbering = @env.numbering
    @footnotes = @env.footnotes
    @footnotes.instance_variable_set(:@counter, 1)
    @converter = Sablon::HTMLConverter.new
  end

  def test_convert_text_inside_div
    input = '<div>Lorem ipsum dolor sit amet</div>'
    expected_output = <<-DOCX.strip
      <w:p>
        <w:pPr><w:pStyle w:val="Normal" /></w:pPr>
        <w:r><w:t xml:space="preserve">Lorem ipsum dolor sit amet</w:t></w:r>
      </w:p>
    DOCX
    assert_equal normalize_wordml(expected_output), process(input)
  end

  def test_convert_text_inside_p
    input = '<p>Lorem ipsum dolor sit amet</p>'
    expected_output = <<-DOCX.strip
      <w:p>
        <w:pPr><w:pStyle w:val="Paragraph" /></w:pPr>
        <w:r><w:t xml:space="preserve">Lorem ipsum dolor sit amet</w:t></w:r>
      </w:p>
    DOCX
    assert_equal normalize_wordml(expected_output), process(input)
  end

  def test_convert_text_inside_multiple_divs
    input = '<div>Lorem ipsum</div><div>dolor sit amet</div>'
    expected_output = <<-DOCX.strip
      <w:p>
        <w:pPr><w:pStyle w:val="Normal" /></w:pPr>
        <w:r><w:t xml:space="preserve">Lorem ipsum</w:t></w:r>
      </w:p>
      <w:p>
        <w:pPr><w:pStyle w:val="Normal" /></w:pPr>
        <w:r><w:t xml:space="preserve">dolor sit amet</w:t></w:r>
      </w:p>
    DOCX
    assert_equal normalize_wordml(expected_output), process(input)
  end

  def test_convert_newline_inside_div
    input = '<div>Lorem ipsum<br>dolor sit amet</div>'
    expected_output = <<-DOCX.strip
      <w:p>
        <w:pPr><w:pStyle w:val="Normal" /></w:pPr>
        <w:r><w:t xml:space="preserve">Lorem ipsum</w:t></w:r>
        <w:r><w:br/></w:r>
        <w:r><w:t xml:space="preserve">dolor sit amet</w:t></w:r>
      </w:p>
    DOCX
    assert_equal normalize_wordml(expected_output), process(input)
  end

  def test_convert_strong_tags_inside_div
    input = '<div>Lorem&nbsp;<strong>ipsum dolor</strong>&nbsp;sit amet</div>'
    expected_output = <<-DOCX.strip
      <w:p>
        <w:pPr><w:pStyle w:val="Normal" /></w:pPr>
        <w:r><w:t xml:space="preserve">Lorem </w:t></w:r>
        <w:r><w:rPr><w:b /></w:rPr><w:t xml:space="preserve">ipsum dolor</w:t></w:r>
        <w:r><w:t xml:space="preserve"> sit amet</w:t></w:r>
      </w:p>
    DOCX
    assert_equal normalize_wordml(expected_output), process(input)
  end

  def test_convert_span_tags_inside_p
    input = '<p>Lorem&nbsp;<span>ipsum dolor</span>&nbsp;sit amet</p>'
    expected_output = <<-DOCX.strip
      <w:p>
        <w:pPr><w:pStyle w:val="Paragraph" /></w:pPr>
        <w:r><w:t xml:space="preserve">Lorem </w:t></w:r>
        <w:r><w:t xml:space="preserve">ipsum dolor</w:t></w:r>
        <w:r><w:t xml:space="preserve"> sit amet</w:t></w:r>
      </w:p>
    DOCX
    assert_equal normalize_wordml(expected_output), process(input)
  end

  def test_convert_u_tags_inside_p
    input = '<p>Lorem&nbsp;<u>ipsum dolor</u>&nbsp;sit amet</p>'
    expected_output = <<-DOCX.strip
      <w:p>
        <w:pPr><w:pStyle w:val="Paragraph" /></w:pPr>
        <w:r><w:t xml:space="preserve">Lorem </w:t></w:r>
        <w:r>
          <w:rPr><w:u w:val="single" /></w:rPr>
          <w:t xml:space="preserve">ipsum dolor</w:t>
        </w:r>
        <w:r><w:t xml:space="preserve"> sit amet</w:t></w:r>
      </w:p>
    DOCX
    assert_equal normalize_wordml(expected_output), process(input)
  end

  def test_convert_em_tags_inside_div
    input = '<div>Lorem&nbsp;<em>ipsum dolor</em>&nbsp;sit amet</div>'
    expected_output = <<-DOCX.strip
      <w:p>
        <w:pPr><w:pStyle w:val="Normal" /></w:pPr>
        <w:r><w:t xml:space="preserve">Lorem </w:t></w:r>
        <w:r><w:rPr><w:i /></w:rPr><w:t xml:space="preserve">ipsum dolor</w:t></w:r>
        <w:r><w:t xml:space="preserve"> sit amet</w:t></w:r>
      </w:p>
    DOCX
    assert_equal normalize_wordml(expected_output), process(input)
  end

  def test_convert_br_tags_inside_strong
    input = '<div><strong><br />Lorem ipsum<br />dolor sit amet</strong></div>'
    expected_output = <<-DOCX
      <w:p>
        <w:pPr><w:pStyle w:val="Normal" /></w:pPr>
        <w:r><w:br/></w:r>
        <w:r>
          <w:rPr><w:b /></w:rPr>
          <w:t xml:space="preserve">Lorem ipsum</w:t></w:r>
          <w:r><w:br/></w:r>
          <w:r>
            <w:rPr><w:b /></w:rPr>
            <w:t xml:space="preserve">dolor sit amet</w:t>
          </w:r>
      </w:p>
    DOCX
    assert_equal normalize_wordml(expected_output), process(input)
  end

  def test_convert_s_tags_inside_p
    input = '<p>Lorem&nbsp;<s>ipsum dolor</s>&nbsp;sit amet</p>'
    expected_output = <<-DOCX.strip
      <w:p>
        <w:pPr><w:pStyle w:val="Paragraph" /></w:pPr>
        <w:r><w:t xml:space="preserve">Lorem </w:t></w:r>
        <w:r>
          <w:rPr><w:strike w:val="true" /></w:rPr>
          <w:t xml:space="preserve">ipsum dolor</w:t>
        </w:r>
        <w:r><w:t xml:space="preserve"> sit amet</w:t></w:r>
      </w:p>
    DOCX
    assert_equal normalize_wordml(expected_output), process(input)
  end

  def test_convert_sub_tags_inside_p
    input = '<p>Lorem&nbsp;<sub>ipsum dolor</sub>&nbsp;sit amet</p>'
    expected_output = <<-DOCX.strip
      <w:p>
        <w:pPr><w:pStyle w:val="Paragraph" /></w:pPr>
        <w:r><w:t xml:space="preserve">Lorem </w:t></w:r>
        <w:r>
          <w:rPr><w:vertAlign w:val="subscript" /></w:rPr>
          <w:t xml:space="preserve">ipsum dolor</w:t>
        </w:r>
        <w:r><w:t xml:space="preserve"> sit amet</w:t></w:r>
      </w:p>
    DOCX
    assert_equal normalize_wordml(expected_output), process(input)
  end

  def test_convert_sup_tags_inside_p
    input = '<p>Lorem&nbsp;<sup>ipsum dolor</sup>&nbsp;sit amet</p>'
    expected_output = <<-DOCX.strip
      <w:p>
        <w:pPr><w:pStyle w:val="Paragraph" /></w:pPr>
        <w:r><w:t xml:space="preserve">Lorem </w:t></w:r>
        <w:r>
          <w:rPr><w:vertAlign w:val="superscript" /></w:rPr>
          <w:t xml:space="preserve">ipsum dolor</w:t>
        </w:r>
        <w:r><w:t xml:space="preserve"> sit amet</w:t></w:r>
      </w:p>
    DOCX
    assert_equal normalize_wordml(expected_output), process(input)
  end

  def test_convert_h1
    input = '<h1>Lorem ipsum dolor</h1>'
    expected_output = <<-DOCX.strip
      <w:p>
        <w:pPr><w:pStyle w:val="Heading1" /></w:pPr>
        <w:r><w:t xml:space="preserve">Lorem ipsum dolor</w:t></w:r>
      </w:p>
    DOCX
    assert_equal normalize_wordml(expected_output), process(input)
  end

  def test_unorderd_lists
    input = '<ul><li>Lorem</li><li>ipsum</li><li>dolor</li></ul>'
    expected_output = <<-DOCX.strip
      <w:p>
        <w:pPr>
          <w:pStyle w:val="ListBullet" />
          <w:numPr>
            <w:ilvl w:val="0" />
            <w:numId w:val="1001" />
          </w:numPr>
        </w:pPr>
        <w:r><w:t xml:space="preserve">Lorem</w:t></w:r>
      </w:p>

      <w:p>
        <w:pPr>
          <w:pStyle w:val="ListBullet" />
          <w:numPr>
            <w:ilvl w:val="0" />
            <w:numId w:val="1001" />
          </w:numPr>
        </w:pPr>
        <w:r><w:t xml:space="preserve">ipsum</w:t></w:r>
      </w:p>

      <w:p>
        <w:pPr>
          <w:pStyle w:val="ListBullet" />
          <w:numPr>
            <w:ilvl w:val="0" />
            <w:numId w:val="1001" />
          </w:numPr>
        </w:pPr>
        <w:r><w:t xml:space="preserve">dolor</w:t></w:r>
      </w:p>
    DOCX
    assert_equal normalize_wordml(expected_output), process(input)

    assert_equal [Sablon::Numbering::Definition.new(1001, 'ListBullet')], @numbering.definitions
  end

  def test_ordered_lists
    input = '<ol><li>Lorem</li><li>ipsum</li><li>dolor</li></ol>'
    expected_output = <<-DOCX.strip
      <w:p>
        <w:pPr>
          <w:pStyle w:val="ListNumber" />
          <w:numPr>
            <w:ilvl w:val="0" />
            <w:numId w:val="1001" />
          </w:numPr>
        </w:pPr>
        <w:r><w:t xml:space="preserve">Lorem</w:t></w:r>
      </w:p>

      <w:p>
        <w:pPr>
          <w:pStyle w:val="ListNumber" />
          <w:numPr>
            <w:ilvl w:val="0" />
            <w:numId w:val="1001" />
          </w:numPr>
        </w:pPr>
        <w:r><w:t xml:space="preserve">ipsum</w:t></w:r>
      </w:p>

      <w:p>
        <w:pPr>
          <w:pStyle w:val="ListNumber" />
          <w:numPr>
            <w:ilvl w:val="0" />
            <w:numId w:val="1001" />
          </w:numPr>
        </w:pPr>
        <w:r><w:t xml:space="preserve">dolor</w:t></w:r>
      </w:p>
    DOCX
    assert_equal normalize_wordml(expected_output), process(input)

    assert_equal [Sablon::Numbering::Definition.new(1001, 'ListNumber')], @numbering.definitions
  end

  def test_mixed_lists
    input = '<ol><li>Lorem</li></ol><ul><li>ipsum</li></ul><ol><li>dolor</li></ol>'
    expected_output = <<-DOCX.strip
      <w:p>
        <w:pPr>
          <w:pStyle w:val="ListNumber" />
          <w:numPr>
            <w:ilvl w:val="0" />
            <w:numId w:val="1001" />
          </w:numPr>
        </w:pPr>
        <w:r><w:t xml:space=\"preserve\">Lorem</w:t></w:r>
      </w:p>

      <w:p>
        <w:pPr>
          <w:pStyle w:val="ListBullet" />
          <w:numPr>
            <w:ilvl w:val="0" />
            <w:numId w:val="1002" />
          </w:numPr>
        </w:pPr>
        <w:r><w:t xml:space="preserve">ipsum</w:t></w:r>
      </w:p>

      <w:p>
        <w:pPr>
          <w:pStyle w:val="ListNumber" />
          <w:numPr>
            <w:ilvl w:val="0" />
            <w:numId w:val="1003" />
          </w:numPr>
        </w:pPr>
        <w:r><w:t xml:space="preserve">dolor</w:t></w:r>
      </w:p>
    DOCX
    assert_equal normalize_wordml(expected_output), process(input)

    assert_equal [Sablon::Numbering::Definition.new(1001, 'ListNumber'),
                  Sablon::Numbering::Definition.new(1002, 'ListBullet'),
                  Sablon::Numbering::Definition.new(1003, 'ListNumber')], @numbering.definitions
  end

  def test_nested_unordered_lists
    input = '<ul><li>Lorem<ul><li>ipsum<ul><li>dolor</li></ul></li></ul></li></ul>'
    expected_output = <<-DOCX.strip
      <w:p>
        <w:pPr>
          <w:pStyle w:val="ListBullet" />
          <w:numPr>
            <w:ilvl w:val="0" />
            <w:numId w:val="1001" />
          </w:numPr>
        </w:pPr>
        <w:r><w:t xml:space="preserve">Lorem</w:t></w:r>
      </w:p>

      <w:p>
        <w:pPr>
          <w:pStyle w:val="ListBullet" />
          <w:numPr>
            <w:ilvl w:val="1" />
            <w:numId w:val="1001" />
          </w:numPr>
        </w:pPr>
        <w:r><w:t xml:space="preserve">ipsum</w:t></w:r>
      </w:p>

      <w:p>
        <w:pPr>
          <w:pStyle w:val="ListBullet" />
          <w:numPr>
            <w:ilvl w:val="2" />
            <w:numId w:val="1001" />
          </w:numPr>
        </w:pPr>
        <w:r><w:t xml:space="preserve">dolor</w:t></w:r>
      </w:p>
    DOCX
    assert_equal normalize_wordml(expected_output), process(input)

    assert_equal [Sablon::Numbering::Definition.new(1001, 'ListBullet')], @numbering.definitions
  end

  def test_basic_html_table_conversion
    input = <<-HTML
      <table>
        <tr><th>TH 1</th><th>TH 2</th></tr>
        <tr><td>Cell 1</td><td>Cell 2</td></tr>
        <tr><td>Cell 3</td><td>Cell 4</td></tr>
      </table>
    HTML
    expected_output = snippet('basic_table')
    #
    assert_equal normalize_wordml(expected_output), process(input)
  end

  def test_footnoteref_conversion
    input = '<p>Lorem<footnoteref id="1"></p>'
    expected_output = <<-DOCX.strip
      <w:p>
        <w:pPr>
          <w:pStyle w:val="Paragraph" />
        </w:pPr>
        <w:r>
          <w:t xml:space="preserve">Lorem</w:t>
        </w:r>
        <w:r>
          <w:rPr>
            <w:rStyle w:val="FootnoteReference" />
          </w:rPr>
          <w:footnoteReference w:id="1"/>
        </w:r>
      </w:p>
    DOCX
    #
    assert_equal normalize_wordml(expected_output), process(input)
  end

  def test_footnote_conversion
    input = '<footnote placeholder="test">Lorem Ipsum</footnote>'
    expected_output = <<-DOCX.strip
      <w:footnote w:id="2">
        <w:p>
          <w:pPr><
            w:pStyle w:val="FootnoteText" />
          </w:pPr>
          <w:r>
            <w:rPr>
              <w:rStyle w:val="FootnoteReference"/>
            </w:rPr>
            <w:footnoteRef/>
          </w:r>
          <w:r>
            <w:t xml:space="preserve">Lorem Ipsum</w:t>
          </w:r>
        </w:p>
      </w:footnote>
    DOCX
    #
    assert_equal '', process(input)
    assert_equal normalize_wordml(expected_output), @footnotes.new_footnotes[0].to_docx
  end

  def test_footnote_and_ref_conversion
    input = '<p>Lorem<footnoteref placeholder="test"></p><footnote placeholder="test">Lorem Ipsum</footnote>'
    expected_output = <<-DOCX.strip
      <w:p>
        <w:pPr>
          <w:pStyle w:val="Paragraph" />
        </w:pPr>
        <w:r>
          <w:t xml:space="preserve">Lorem</w:t>
        </w:r>
        <w:r>
          <w:rPr>
            <w:rStyle w:val="FootnoteReference" />
          </w:rPr>
          <w:footnoteReference w:id="2"/>
        </w:r>
      </w:p>
    DOCX
    #
    assert_equal normalize_wordml(expected_output), process(input)
  end

  def test_insertion_field_conversion
    input = '<p><ins placeholder="{date}">DATE \\@ "yyyy-dd-MM hh:mm:ss"</ins></p>'
    expected_output = <<-DOCX.strip
      <w:p>
        <w:pPr>
          <w:pStyle w:val="Paragraph" />
        </w:pPr>
        <w:r>
          <w:rPr><w:noProof /></w:rPr>
          <w:fldChar w:fldCharType="begin"/>
        </w:r>
        <w:r>
          <w:rPr><w:noProof /></w:rPr>
          <w:instrText xml:space="preserve"> DATE \\@ "yyyy-dd-MM hh:mm:ss" </w:instrText>
        </w:r>
        <w:r>
          <w:rPr><w:noProof /></w:rPr>
          <w:fldChar w:fldCharType="separate"/>
        </w:r>
        <w:r>
          <w:rPr><w:noProof /></w:rPr>
          <w:t xml:space="preserve">{date}</w:t>
        </w:r>
        <w:r>
          <w:rPr><w:noProof /></w:rPr>
          <w:fldChar w:fldCharType="end"/>
        </w:r>
      </w:p>
    DOCX
    #
    assert_equal normalize_wordml(expected_output), process(input)
  end

  def test_bookmark_conversion
    input = '<p><bookmark name="test">Lorem</bookmark> Ipsum</p>'
    expected_output = <<-DOCX.strip
      <w:p>
        <w:pPr>
          <w:pStyle w:val="Paragraph" />
        </w:pPr>
        <w:bookmarkStart w:id="1" w:name="test"/>
        <w:r>
          <w:t xml:space="preserve">Lorem</w:t>
        </w:r>
        <w:bookmarkEnd w:id="1"/>
        <w:r>
          <w:t xml:space="preserve"> Ipsum</w:t>
        </w:r>
      </w:p>
    DOCX
    #
    assert_equal normalize_wordml(expected_output), process(input)
  end

  def test_nonunique_bookmark_name
    input = '<p><bookmark name="test">Lorem</bookmark> <bookmark name="test">Ipsum</bookmark></p>'
    e = assert_raises Sablon::ContextError do
      process(input)
    end
    assert_equal "Bookmark name already in use: test", e.message
  end

  def test_caption_conversion
    input = '<caption type="table" name="tab-data">Lorem Ipsum</caption>'
    expected_output = <<-DOCX.strip
      <w:p>
        <w:pPr>
          <w:pStyle w:val="Caption" />
        </w:pPr>
        <w:bookmarkStart w:id="1" w:name="tab-data"/>
        <w:r>
          <w:t xml:space="preserve">Table</w:t>
        </w:r>
        <w:r>
          <w:fldChar w:fldCharType="begin"/>
        </w:r>
        <w:r>
          <w:instrText xml:space="preserve"> SEQ Table \\# " # " </w:instrText>
        </w:r>
        <w:r>
          <w:fldChar w:fldCharType="separate"/>
        </w:r>
        <w:r>
          <w:t xml:space="preserve"> # </w:t>
        </w:r>
        <w:r>
          <w:fldChar w:fldCharType="end"/>
        </w:r>
        <w:bookmarkEnd w:id="1"/>
        <w:r>
          <w:t xml:space="preserve">Lorem Ipsum</w:t>
        </w:r>
      </w:p>
    DOCX
    #
    assert_equal normalize_wordml(expected_output), process(input)
  end

  private

  def process(input)
    @converter.process(input, @env)
  end

  def normalize_wordml(wordml)
    wordml.gsub(/^\s+/, '').tr("\n", '')
  end
end

class HTMLConverterStyleTest < Sablon::TestCase
  def setup
    super
    @env = Sablon::Environment.new(nil)
    @converter = Sablon::HTMLConverter.new
  end

  # testing direct CSS style -> WordML conversion for paragraphs

  def test_paragraph_with_background_color
    input = '<p style="background-color: #123456"></p>'
    expected_output = para_with_ppr('<w:shd w:val="clear" w:fill="123456" />')
    assert_equal normalize_wordml(expected_output), process(input)
  end

  def test_paragraph_with_borders
    input = '<p style="border: 1px"></p>'
    ppr = <<-DOCX.gsub(/^\s*/, '').delete("\n")
      <w:pBdr>
        <w:top w:sz="2" w:val="single" w:color="000000" />
        <w:bottom w:sz="2" w:val="single" w:color="000000" />
        <w:left w:sz="2" w:val="single" w:color="000000" />
        <w:right w:sz="2" w:val="single" w:color="000000" />
      </w:pBdr>
    DOCX
    expected_output = para_with_ppr(ppr)
    assert_equal normalize_wordml(expected_output), process(input)
    #
    input = '<p style="border: 1px wavy"></p>'
    ppr = <<-DOCX.gsub(/^\s*/, '').delete("\n")
      <w:pBdr>
        <w:top w:sz="2" w:val="wavy" w:color="000000" />
        <w:bottom w:sz="2" w:val="wavy" w:color="000000" />
        <w:left w:sz="2" w:val="wavy" w:color="000000" />
        <w:right w:sz="2" w:val="wavy" w:color="000000" />
      </w:pBdr>
    DOCX
    expected_output = para_with_ppr(ppr)
    assert_equal normalize_wordml(expected_output), process(input)
    #
    input = '<p style="border: 1px wavy #123456"></p>'
    ppr = <<-DOCX.gsub(/^\s*/, '').delete("\n")
      <w:pBdr>
        <w:top w:sz="2" w:val="wavy" w:color="123456" />
        <w:bottom w:sz="2" w:val="wavy" w:color="123456" />
        <w:left w:sz="2" w:val="wavy" w:color="123456" />
        <w:right w:sz="2" w:val="wavy" w:color="123456" />
      </w:pBdr>
    DOCX
    expected_output = para_with_ppr(ppr)
    assert_equal normalize_wordml(expected_output), process(input)
  end

  def test_paragraph_with_text_align
    input = '<p style="text-align: both"></p>'
    expected_output = para_with_ppr('<w:jc w:val="both" />')
    assert_equal normalize_wordml(expected_output), process(input)
  end

  def test_paragraph_with_vertical_align
    input = '<p style="vertical-align: baseline"></p>'
    expected_output = para_with_ppr('<w:textAlignment w:val="baseline" />')
    assert_equal normalize_wordml(expected_output), process(input)
  end


  def test_paragraph_with_unsupported_property
    input = '<p style="unsupported: true"></p>'
    expected_output = para_with_ppr('')
    assert_equal normalize_wordml(expected_output), process(input)
  end

  def test_run_with_background_color
    input = '<p><span style="background-color: #123456">test</span></p>'
    expected_output = run_with_rpr('<w:shd w:val="clear" w:fill="123456" />')
    assert_equal normalize_wordml(expected_output), process(input)
  end

  def test_run_with_color
    input = '<p><span style="color: #123456">test</span></p>'
    expected_output = run_with_rpr('<w:color w:val="123456" />')
    assert_equal normalize_wordml(expected_output), process(input)
  end

  def test_run_with_font_size
    input = '<p><span style="font-size: 20">test</span></p>'
    expected_output = run_with_rpr('<w:sz w:val="40" />')
    assert_equal normalize_wordml(expected_output), process(input)

    # test that non-numeric are ignored
    input = '<p><span style="font-size: 20pts">test</span></p>'
    assert_equal normalize_wordml(expected_output), process(input)

    # test that floats round up
    input = '<p><span style="font-size: 19.1pts">test</span></p>'
    assert_equal normalize_wordml(expected_output), process(input)
  end

  def test_run_with_font_style
    input = '<p><span style="font-style: bold">test</span></p>'
    expected_output = run_with_rpr('<w:b />')
    assert_equal normalize_wordml(expected_output), process(input)

    # test that non-numeric are ignored
    input = '<p><span style="font-style: italic">test</span></p>'
    expected_output = run_with_rpr('<w:i />')
    assert_equal normalize_wordml(expected_output), process(input)
  end

  def test_run_with_font_wieght
    input = '<p><span style="font-weight: bold">test</span></p>'
    expected_output = run_with_rpr('<w:b />')
    assert_equal normalize_wordml(expected_output), process(input)
  end

  def test_run_with_text_decoration
    # testing underline configurations
    input = '<p><span style="text-decoration: underline">test</span></p>'
    expected_output = run_with_rpr('<w:u w:val="single" />')
    assert_equal normalize_wordml(expected_output), process(input)

    input = '<p><span style="text-decoration: underline dash">test</span></p>'
    expected_output = run_with_rpr('<w:u w:val="dash" w:color="auto" />')
    assert_equal normalize_wordml(expected_output), process(input)

    input = '<p><span style="text-decoration: underline dash #123456">test</span></p>'
    expected_output = run_with_rpr('<w:u w:val="dash" w:color="123456" />')
    assert_equal normalize_wordml(expected_output), process(input)

    # testing line-through
    input = '<p><span style="text-decoration: line-through">test</span></p>'
    expected_output = run_with_rpr('<w:strike w:val="true" />')
    assert_equal normalize_wordml(expected_output), process(input)

    # testing that unsupported values are passed through as a toggle
    input = '<p><span style="text-decoration: strike">test</span></p>'
    expected_output = run_with_rpr('<w:strike w:val="true" />')
    assert_equal normalize_wordml(expected_output), process(input)

    input = '<p><span style="text-decoration: emboss">test</span></p>'
    expected_output = run_with_rpr('<w:emboss w:val="true" />')
    assert_equal normalize_wordml(expected_output), process(input)
  end

  def test_run_with_vertical_align
    input = '<p><span style="vertical-align: subscript">test</span></p>'
    expected_output = run_with_rpr('<w:vertAlign w:val="subscript" />')
    assert_equal normalize_wordml(expected_output), process(input)

    input = '<p><span style="vertical-align: superscript">test</span></p>'
    expected_output = run_with_rpr('<w:vertAlign w:val="superscript" />')
    assert_equal normalize_wordml(expected_output), process(input)
  end

  def test_run_with_unsupported_property
    input = '<p><span style="unsupported: true">test</span></p>'
    expected_output = '<w:p><w:pPr><w:pStyle w:val="Paragraph" /></w:pPr><w:r><w:t xml:space="preserve">test</w:t></w:r></w:p>'
    assert_equal normalize_wordml(expected_output), process(input)
  end

  # tests with nested runs and styles

  def test_paragraph_props_passed_to_runs
    input = '<p style="color: #123456"><b>Lorem</b><span>ipsum</span></p>'
    expected_output = <<-DOCX.strip
      <w:p>
        <w:pPr>
          <w:pStyle w:val="Paragraph" />
        </w:pPr>
        <w:r>
          <w:rPr>
             <w:color w:val="123456" />
            <w:b />
          </w:rPr>
          <w:t xml:space="preserve">Lorem</w:t>
        </w:r>
        <w:r>
          <w:rPr>
            <w:color w:val="123456" />
          </w:rPr>
          <w:t xml:space="preserve">ipsum</w:t>
        </w:r>
      </w:p>
    DOCX
    assert_equal normalize_wordml(expected_output), process(input)
  end

  def test_run_prop_override_paragraph_prop
    input = '<p style="text-align: center; color: #FF0000">Lorem<span style="color: blue;">ipsum</span></p>'
    expected_output = <<-DOCX.strip
      <w:p>
        <w:pPr>
          <w:jc w:val="center" />
          <w:pStyle w:val="Paragraph" />
        </w:pPr>
        <w:r>
          <w:rPr>
            <w:color w:val="FF0000" />
          </w:rPr>
          <w:t xml:space="preserve">Lorem</w:t>
        </w:r>
        <w:r>
          <w:rPr>
            <w:color w:val="blue" />
          </w:rPr>
          <w:t xml:space="preserve">ipsum</w:t>
        </w:r>
      </w:p>
    DOCX
    assert_equal normalize_wordml(expected_output), process(input)
  end

  def test_table_with_borders
    input = '<table style="border: 1px"><tr><td>Cell 1</td></tr></table>'
    tblpr = <<-DOCX.strip
      <w:tblBorders>
        <w:top w:sz="2" w:val="single" w:color="000000" />
        <w:start w:sz="2" w:val="single" w:color="000000" />
        <w:bottom w:sz="2" w:val="single" w:color="000000" />
        <w:end w:sz="2" w:val="single" w:color="000000" />
        <w:insideH w:sz="2" w:val="single" w:color="000000" />
        <w:insideV w:sz="2" w:val="single" w:color="000000" />
      </w:tblBorders>
    DOCX
    expected_output = table_with_style(tblPr: tblpr)
    assert_equal normalize_wordml(expected_output), process(input)
  end

  def test_table_with_margins
    # single value test
    input = '<table style="margin: 1px"><tr><td>Cell 1</td></tr></table>'
    tblpr = <<-DOCX.strip
      <w:tblCellMar>
        <w:top w:w="2" w:type="dxa" />
        <w:end w:w="2" w:type="dxa" />
        <w:bottom w:w="2" w:type="dxa" />
        <w:start w:w="2" w:type="dxa" />
      </w:tblCellMar>
    DOCX
    expected_output = table_with_style(tblPr: tblpr)
    assert_equal normalize_wordml(expected_output), process(input)
    # double value test
    input = '<table style="margin: 1px 2px"><tr><td>Cell 1</td></tr></table>'
    tblpr = <<-DOCX.strip
      <w:tblCellMar>
        <w:top w:w="2" w:type="dxa" />
        <w:end w:w="4" w:type="dxa" />
        <w:bottom w:w="2" w:type="dxa" />
        <w:start w:w="4" w:type="dxa" />
      </w:tblCellMar>
    DOCX
    expected_output = table_with_style(tblPr: tblpr)
    assert_equal normalize_wordml(expected_output), process(input)
    # triple value test
    input = '<table style="margin: 1px 2px 3px"><tr><td>Cell 1</td></tr></table>'
    tblpr = <<-DOCX.strip
      <w:tblCellMar>
        <w:top w:w="2" w:type="dxa" />
        <w:end w:w="4" w:type="dxa" />
        <w:bottom w:w="6" w:type="dxa" />
        <w:start w:w="4" w:type="dxa" />
      </w:tblCellMar>
    DOCX
    expected_output = table_with_style(tblPr: tblpr)
    assert_equal normalize_wordml(expected_output), process(input)
    # four values test
    input = '<table style="margin: 1px 2px 3px 4px"><tr><td>Cell 1</td></tr></table>'
    tblpr = <<-DOCX.strip
      <w:tblCellMar>
        <w:top w:w="2" w:type="dxa" />
        <w:end w:w="4" w:type="dxa" />
        <w:bottom w:w="6" w:type="dxa" />
        <w:start w:w="8" w:type="dxa" />
      </w:tblCellMar>
    DOCX
    expected_output = table_with_style(tblPr: tblpr)
    assert_equal normalize_wordml(expected_output), process(input)
  end

  def test_table_with_cellspacing
    input = '<table style="cellspacing: 10"><tr><td>Cell 1</td></tr></table>'
    expected_output = table_with_style(tblPr: '<w:tblCellSpacing w:w="20" w:type="dxa" />')
    assert_equal normalize_wordml(expected_output), process(input)
  end

  def test_table_with_text_align
    input = '<table style="text-align: center"><tr><td>Cell 1</td></tr></table>'
    expected_output = table_with_style(tblPr: '<w:jc w:val="center" />')
    assert_equal normalize_wordml(expected_output), process(input)
  end

  def test_table_with_width
    input = '<table style="width: 1000"><tr><td>Cell 1</td></tr></table>'
    expected_output = table_with_style(tblPr: '<w:tblW w:w="2000" w:type="dxa" />')
    assert_equal normalize_wordml(expected_output), process(input)
  end

  def test_table_with_row_text_align
    input = '<table><tr style="text-align: center"><td>Cell 1</td></tr></table>'
    expected_output = table_with_style(trPr: '<w:jc w:val="center" />')
    assert_equal normalize_wordml(expected_output), process(input)
  end

  def test_table_with_cell_borders
    input = '<table><tr><td style="border: 1px">Cell 1</td></tr></table>'
    tcpr = <<-DOCX.strip
      <w:tcBorders>
        <w:top w:sz="2" w:val="single" w:color="000000" />
        <w:start w:sz="2" w:val="single" w:color="000000" />
        <w:bottom w:sz="2" w:val="single" w:color="000000" />
        <w:end w:sz="2" w:val="single" w:color="000000" />
        <w:insideH w:sz="2" w:val="single" w:color="000000" />
        <w:insideV w:sz="2" w:val="single" w:color="000000" />
      </w:tcBorders>
    DOCX
    expected_output = table_with_style(tcPr: tcpr)
    assert_equal normalize_wordml(expected_output), process(input)
    # test that the proeprty will be passed onto cells from rows
    input = '<table><tr style="border: 1px"><td>Cell 1</td></tr></table>'
    tcpr = <<-DOCX.strip
      <w:tcBorders>
        <w:top w:sz="2" w:val="single" w:color="000000" />
        <w:start w:sz="2" w:val="single" w:color="000000" />
        <w:bottom w:sz="2" w:val="single" w:color="000000" />
        <w:end w:sz="2" w:val="single" w:color="000000" />
        <w:insideH w:sz="2" w:val="single" w:color="000000" />
        <w:insideV w:sz="2" w:val="single" w:color="000000" />
      </w:tcBorders>
    DOCX
    expected_output = table_with_style(tcPr: tcpr)
    assert_equal normalize_wordml(expected_output), process(input)
  end

  def test_table_with_cell_colspan
    input = '<table><tr><td style="colspan: 2">Cell 1</td></tr></table>'
    expected_output = table_with_style(tcPr: '<w:gridSpan w:val="2" />')
    assert_equal normalize_wordml(expected_output), process(input)
  end

  def test_table_with_cell_margin
    input = '<table><tr><td style="margin: 1px">Cell 1</td></tr></table>'
    tcpr = <<-DOCX.strip
      <w:tcMar>
        <w:top w:w="2" w:type="dxa" />
        <w:end w:w="2" w:type="dxa" />
        <w:bottom w:w="2" w:type="dxa" />
        <w:start w:w="2" w:type="dxa" />
      </w:tcMar>
    DOCX
    expected_output = table_with_style(tcPr: tcpr)
    assert_equal normalize_wordml(expected_output), process(input)
  end

  def test_table_with_cell_rowspan
    # test start
    input = '<table><tr><td style="rowspan: start">Cell 1</td></tr></table>'
    expected_output = table_with_style(tcPr: '<w:vMerge w:val="restart" />')
    assert_equal normalize_wordml(expected_output), process(input)
    # test continue
    input = '<table><tr><td style="rowspan: continue">Cell 1</td></tr></table>'
    expected_output = table_with_style(tcPr: '<w:vMerge w:val="continue" />')
    assert_equal normalize_wordml(expected_output), process(input)
    # test end
    input = '<table><tr><td style="rowspan: end">Cell 1</td></tr></table>'
    expected_output = table_with_style(tcPr: '<w:vMerge />')
    assert_equal normalize_wordml(expected_output), process(input)
  end

  def test_table_with_cell_vertical_align
    input = '<table><tr><td style="vertical-align: top">Cell 1</td></tr></table>'
    expected_output = table_with_style(tcPr: '<w:vAlign w:val="top" />')
    assert_equal normalize_wordml(expected_output), process(input)
  end

  def test_table_with_cell_white_space
    # test nowrap
    input = '<table><tr><td style="white-space: nowrap">Cell 1</td></tr></table>'
    expected_output = table_with_style(tcPr: '<w:noWrap />')
    assert_equal normalize_wordml(expected_output), process(input)
    # test fit
    input = '<table><tr><td style="white-space: fit">Cell 1</td></tr></table>'
    expected_output = table_with_style(tcPr: '<w:tcFitText w:val="true" />')
    assert_equal normalize_wordml(expected_output), process(input)
  end

  def test_table_with_cell_width
    input = '<table><tr><td style="width: 1000">Cell 1</td></tr></table>'
    expected_output = table_with_style(tcPr: '<w:tcW w:w="2000" w:type="dxa" />')
    assert_equal normalize_wordml(expected_output), process(input)
  end

  private

  def process(input)
    @converter.process(input, @env)
  end

  def para_with_ppr(ppr_str)
    para_str = '<w:p><w:pPr>%s<w:pStyle w:val="Paragraph" /></w:pPr></w:p>'
    format(para_str, ppr_str)
  end

  def run_with_rpr(rpr_str)
    para_str = <<-DOCX.strip
      <w:p>
        <w:pPr>
          <w:pStyle w:val="Paragraph" />
        </w:pPr>
        <w:r>
          <w:rPr>
            %s
          </w:rPr>
          <w:t xml:space="preserve">test</w:t>
        </w:r>
      </w:p>
    DOCX
    format(para_str, rpr_str)
  end

  def table_with_style(properties)
    properties = { tblPr: nil, trPr: nil, tcPr: nil,
                   pPr: nil, rPr: nil }.merge(properties)
    properties = properties.map do |key, value|
      "<w:#{key}>#{value}</w:#{key}>" if value
    end
    #
    table_str = <<-DOCX.strip
      <w:tbl>
        %s
        <w:tr>
          %s
          <w:tc>
            %s
            <w:p>
              %s
              <w:r>
                %s
                <w:t xml:space="preserve">Cell 1</w:t>
              </w:r>
            </w:p>
          </w:tc>
        </w:tr>
      </w:tbl>
    DOCX
    #
    format(table_str, *properties)
  end

  def normalize_wordml(wordml)
    wordml.gsub(/^\s+/, '').tr("\n", '')
  end
end
