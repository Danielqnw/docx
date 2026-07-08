# frozen_string_literal: true

require 'spec_helper'
require 'docx'
require 'docx/containers'

describe Docx::Elements::Containers::Paragraph do
  NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

  def paragraph_from_xml(xml)
    doc = Nokogiri::XML(
      %(<w:p xmlns:w="#{NS}" xmlns:xml="http://www.w3.org/XML/1998/namespace">#{xml}</w:p>)
    )
    described_class.new(doc.root)
  end

  def paragraph_text(p)
    p.node.xpath('.//w:t').map(&:content).join
  end

  describe '#substitute' do
    it 'replaces a placeholder split across two runs and preserves first run formatting' do
      p = paragraph_from_xml(<<~XML)
        <w:r><w:rPr><w:b/></w:rPr><w:t>{{ti</w:t></w:r>
        <w:r><w:t>tle}}</w:t></w:r>
      XML

      p.substitute('{{title}}', 'Hello')

      expect(paragraph_text(p)).to eq('Hello')
      expect(p.node.at_xpath('w:r[1]/w:rPr/w:b')).not_to be_nil
    end

    it 'replaces within a single run' do
      p = paragraph_from_xml('<w:r><w:t>Hello {{name}}!</w:t></w:r>')

      p.substitute('{{name}}', 'World')

      expect(paragraph_text(p)).to eq('Hello World!')
    end

    it 'replaces multiple occurrences in the same paragraph' do
      p = paragraph_from_xml('<w:r><w:t>{{x}} and {{x}}</w:t></w:r>')

      p.substitute('{{x}}', 'ok')

      expect(paragraph_text(p)).to eq('ok and ok')
    end

    it 'supports regexp with capture group replacement' do
      p = paragraph_from_xml('<w:r><w:t>ID: {{user-42}}</w:t></w:r>')

      p.substitute(/\{\{user-(\d+)\}\}/, 'User #\1')

      expect(paragraph_text(p)).to eq('ID: User #42')
    end

    it 'does not modify the paragraph when there is no match' do
      xml = '<w:r><w:t>unchanged</w:t></w:r>'
      p = paragraph_from_xml(xml)
      original = p.node.to_xml

      expect { p.substitute('{{missing}}', 'x') }.not_to raise_error
      expect(p.node.to_xml).to eq(original)
    end

    context 'with multiline: true' do
      it 'inserts w:br nodes for newlines in replacement' do
        p = paragraph_from_xml('<w:r><w:t>{{body}}</w:t></w:r>')

        p.substitute('{{body}}', "Line1\nLine2", multiline: true)

        expect(p.node.xpath('.//w:br').size).to eq(1)
        expect(paragraph_text(p)).to eq("Line1Line2")
        t_nodes = p.node.xpath('.//w:t')
        expect(t_nodes[0].content).to eq('Line1')
        expect(t_nodes[1].content).to eq('Line2')
      end

      it 'does not insert w:br when multiline is false' do
        p = paragraph_from_xml('<w:r><w:t>{{body}}</w:t></w:r>')

        p.substitute('{{body}}', "Line1\nLine2", multiline: false)

        expect(p.node.xpath('.//w:br')).to be_empty
        expect(paragraph_text(p)).to eq("Line1\nLine2")
      end
    end
  end
end
