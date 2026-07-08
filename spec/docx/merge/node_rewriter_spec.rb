# frozen_string_literal: true

require 'spec_helper'
require 'docx'

describe Docx::Merge::NodeRewriter do
  def xml_ns
    Docx::Document::XML_NAMESPACES
  end

  def w_ns
    xml_ns['w']
  end

  def frag(xml)
    Nokogiri::XML(xml).root
  end

  def rewriter(style_id_map: {}, num_id_map: {})
    described_class.new(style_id_map: style_id_map, num_id_map: num_id_map)
  end

  describe '#rewrite' do
    it 'rewrites table style references' do
      node = frag(<<~XML)
        <w:tbl xmlns:w="#{w_ns}">
          <w:tblPr>
            <w:tblStyle w:val="TblX"/>
          </w:tblPr>
          <w:tr>
            <w:tc>
              <w:p/>
            </w:tc>
          </w:tr>
        </w:tbl>
      XML

      rewriter(style_id_map: { 'TblX' => 'm1_TblX' }).rewrite(node)

      expect(node.at_xpath('.//w:tblStyle', xml_ns)['w:val']).to eq('m1_TblX')
    end

    it 'rewrites deeply nested paragraph and run style references' do
      node = frag(<<~XML)
        <w:tbl xmlns:w="#{w_ns}">
          <w:tr>
            <w:tc>
              <w:p>
                <w:pPr>
                  <w:pStyle w:val="Heading1"/>
                </w:pPr>
                <w:r>
                  <w:rPr>
                    <w:rStyle w:val="Emphasis"/>
                  </w:rPr>
                  <w:t>text</w:t>
                </w:r>
              </w:p>
            </w:tc>
          </w:tr>
        </w:tbl>
      XML

      rewriter(style_id_map: { 'Heading1' => 'm1_Heading1', 'Emphasis' => 'm1_Emphasis' }).rewrite(node)

      expect(node.at_xpath('.//w:pStyle', xml_ns)['w:val']).to eq('m1_Heading1')
      expect(node.at_xpath('.//w:rStyle', xml_ns)['w:val']).to eq('m1_Emphasis')
    end

    it 'rewrites numId only under numPr' do
      node = frag(<<~XML)
        <w:p xmlns:w="#{w_ns}">
          <w:pPr>
            <w:numPr>
              <w:numId w:val="42"/>
            </w:numPr>
          </w:pPr>
          <w:num w:numId="99"/>
        </w:p>
      XML

      rewriter(num_id_map: { '42' => '7' }).rewrite(node)

      expect(node.at_xpath('.//w:numPr/w:numId', xml_ns)['w:val']).to eq('7')
      expect(node.at_xpath('./w:num', xml_ns)['w:numId']).to eq('99')
    end

    it 'leaves unmapped style and numId values unchanged' do
      node = frag(<<~XML)
        <w:p xmlns:w="#{w_ns}">
          <w:pPr>
            <w:pStyle w:val="KeepMe"/>
            <w:numPr>
              <w:numId w:val="100"/>
            </w:numPr>
          </w:pPr>
          <w:r>
            <w:rPr>
              <w:rStyle w:val="AlsoKeep"/>
            </w:rPr>
          </w:r>
        </w:p>
      XML

      rewriter(style_id_map: { 'Other' => 'mapped' }, num_id_map: { '200' => '300' }).rewrite(node)

      expect(node.at_xpath('.//w:pStyle', xml_ns)['w:val']).to eq('KeepMe')
      expect(node.at_xpath('.//w:rStyle', xml_ns)['w:val']).to eq('AlsoKeep')
      expect(node.at_xpath('.//w:numPr/w:numId', xml_ns)['w:val']).to eq('100')
    end

    it 'is a no-op when maps are empty' do
      xml = <<~XML
        <w:p xmlns:w="#{w_ns}">
          <w:pPr>
            <w:pStyle w:val="Heading1"/>
          </w:pPr>
        </w:p>
      XML
      node = frag(xml)

      rewriter.rewrite(node)

      expect(node.to_xml).to eq(frag(xml).to_xml)
    end

    it 'preserves w:val namespace prefix in serialized output' do
      node = frag(<<~XML)
        <w:p xmlns:w="#{w_ns}">
          <w:pPr>
            <w:pStyle w:val="Heading1"/>
          </w:pPr>
        </w:p>
      XML

      rewriter(style_id_map: { 'Heading1' => 'm1_Heading1' }).rewrite(node)
      serialized = node.to_xml

      expect(serialized).to include('w:val="m1_Heading1"')
      expect(serialized.scan('xmlns:w=').length).to eq(1)
    end

    it 'returns the same node object' do
      node = frag(%(<w:p xmlns:w="#{w_ns}"><w:pPr><w:pStyle w:val="A"/></w:pPr></w:p>))

      result = rewriter(style_id_map: { 'A' => 'B' }).rewrite(node)

      expect(result).to equal(node)
    end

    it 'rewrites style references when the root node itself is a style element' do
      node = frag(%(<w:pStyle xmlns:w="#{w_ns}" w:val="Heading1"/>))

      rewriter(style_id_map: { 'Heading1' => 'm1_Heading1' }).rewrite(node)

      expect(node['w:val']).to eq('m1_Heading1')
    end

    it 'rewrites numId when the root node is numPr' do
      node = frag(<<~XML)
        <w:numPr xmlns:w="#{w_ns}">
          <w:numId w:val="42"/>
        </w:numPr>
      XML

      rewriter(num_id_map: { '42' => '7' }).rewrite(node)

      expect(node.at_xpath('./w:numId', xml_ns)['w:val']).to eq('7')
    end
  end
end
