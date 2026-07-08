# frozen_string_literal: true

require 'spec_helper'
require 'docx'

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
W14_NS = 'http://schemas.microsoft.com/office/word/2010/wordml'

describe 'Docx::Document checkbox methods' do
  def build_document(xml)
    doc = Docx::Document.allocate
    doc.instance_variable_set(:@doc, Nokogiri::XML(xml))
    doc
  end

  def paragraph_text(doc, index)
    doc.instance_variable_get(:@doc).xpath('//w:p', 'w' => W_NS)[index].xpath('.//w:t', 'w' => W_NS).map(&:text).join
  end

  describe '#set_checkbox' do
    let(:char_checkbox_xml) do
      <<~XML
        <w:document xmlns:w="#{W_NS}">
          <w:body>
            <w:p><w:r><w:t>\u2610 选项A</w:t></w:r></w:p>
            <w:p><w:r><w:t>\u2610 其他</w:t></w:r></w:p>
          </w:body>
        </w:document>
      XML
    end

    it 'checks a character checkbox in the matching paragraph only' do
      doc = build_document(char_checkbox_xml)

      expect(doc.set_checkbox('选项A', checked: true)).to eq(true)
      expect(paragraph_text(doc, 0)).to eq("\u2611 选项A")
      expect(paragraph_text(doc, 1)).to eq("\u2610 其他")
    end

    it 'is idempotent when already checked' do
      doc = build_document(char_checkbox_xml)
      doc.set_checkbox('选项A', checked: true)

      expect(doc.set_checkbox('选项A', checked: true)).to eq(true)
      expect(paragraph_text(doc, 0)).to eq("\u2611 选项A")
    end

    it 'unchecks a character checkbox' do
      doc = build_document(char_checkbox_xml)
      doc.set_checkbox('选项A', checked: true)

      expect(doc.set_checkbox('选项A', checked: false)).to eq(true)
      expect(paragraph_text(doc, 0)).to eq("\u2610 选项A")
    end

    it 'supports custom checked and unchecked glyphs' do
      custom_xml = <<~XML
        <w:document xmlns:w="#{W_NS}">
          <w:body>
            <w:p><w:r><w:t>[ ] 自定义</w:t></w:r></w:p>
          </w:body>
        </w:document>
      XML
      doc = build_document(custom_xml)

      doc.set_checkbox('[ ]', checked: true, unchecked_glyph: '[ ]', checked_glyph: '[x]')

      expect(paragraph_text(doc, 0)).to eq('[x] 自定义')
    end

    it 'raises ImagePlaceholderNotFound when locator is missing' do
      doc = build_document(char_checkbox_xml)

      expect { doc.set_checkbox('missing', checked: true) }
        .to raise_error(Docx::Errors::ImagePlaceholderNotFound, 'Checkbox locator not found: missing')
    end
  end

  describe '#check_content_control' do
    def content_control_xml(tag: 'opt_a', alias_val: 'opt_a')
      <<~XML
        <w:document xmlns:w="#{W_NS}" xmlns:w14="#{W14_NS}">
          <w:body>
            <w:p>
              <w:sdt>
                <w:sdtPr>
                  <w:alias w:val="#{alias_val}"/>
                  <w:tag w:val="#{tag}"/>
                  <w14:checkbox>
                    <w14:checked w14:val="0"/>
                    <w14:checkedState w14:val="2612" w14:font="MS Gothic"/>
                    <w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/>
                  </w14:checkbox>
                </w:sdtPr>
                <w:sdtContent><w:r><w:t>\u2610</w:t></w:r></w:sdtContent>
              </w:sdt>
            </w:p>
          </w:body>
        </w:document>
      XML
    end

    it 'checks a content control checkbox by tag and syncs the display glyph' do
      doc = build_document(content_control_xml)
      xml_doc = doc.instance_variable_get(:@doc)

      expect(doc.check_content_control('opt_a', checked: true)).to eq(true)
      expect(xml_doc.at_xpath('//w14:checked', 'w14' => W14_NS)['w14:val']).to eq('1')
      expect(xml_doc.at_xpath('//w:sdtContent//w:t', 'w' => W_NS).text).to eq("\u2612")
    end

    it 'matches a content control checkbox by alias' do
      doc = build_document(content_control_xml(tag: 'tag_only', alias_val: 'alias_match'))
      xml_doc = doc.instance_variable_get(:@doc)

      expect(doc.check_content_control('alias_match', checked: true)).to eq(true)
      expect(xml_doc.at_xpath('//w14:checked', 'w14' => W14_NS)['w14:val']).to eq('1')
    end

    it 'raises ImagePlaceholderNotFound when tag or alias is missing' do
      doc = build_document(content_control_xml)

      expect { doc.check_content_control('missing', checked: true) }
        .to raise_error(Docx::Errors::ImagePlaceholderNotFound, 'Content control checkbox not found: missing')
    end
  end
end
