require 'spec_helper'
require 'docx'
require 'stringio'

RSpec.describe Docx::Document, '#replace_image_by_placeholder' do
  let(:rels_xml) do
    <<~XML
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship
          Id="rId5"
          Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
          Target="media/image1.png" />
      </Relationships>
    XML
  end
  let(:zip) { instance_double(Zip::File) }

  def fake_png(w, h) = "\x89PNG\r\n\x1A\n".b + "\x00\x00\x00\rIHDR".b + [w, h].pack('N2') + "\x08\x02\x00\x00\x00".b

  def build_doc(doc_xml)
    d = Docx::Document.allocate
    d.instance_variable_set(:@replace, {})
    d.instance_variable_set(:@doc, Nokogiri::XML(doc_xml))
    d.instance_variable_set(:@rels, Nokogiri::XML(rels_xml))
    d.instance_variable_set(:@zip, zip)
    d
  end

  describe 'in a regular paragraph (non-table)' do
    let(:paragraph_doc_xml) do
      <<~XML
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                    xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
          <w:body>
            <w:p>
              <w:r><w:t>{{photo}}</w:t></w:r>
              <w:r>
                <w:drawing>
                  <wp:inline>
                    <wp:extent cx="1600" cy="900" />
                    <a:graphic>
                      <a:graphicData>
                        <pic:pic>
                          <pic:blipFill>
                            <a:blip r:embed="rId5" />
                            <a:stretch><a:fillRect/></a:stretch>
                          </pic:blipFill>
                          <pic:spPr>
                            <a:xfrm>
                              <a:ext cx="1600" cy="900" />
                            </a:xfrm>
                          </pic:spPr>
                        </pic:pic>
                      </a:graphicData>
                    </a:graphic>
                  </wp:inline>
                </w:drawing>
              </w:r>
            </w:p>
          </w:body>
        </w:document>
      XML
    end
    let(:doc) { build_doc(paragraph_doc_xml) }

    before { allow(zip).to receive(:find_entry).with('word/media/image1.png').and_return(true) }

    it 'replaces image bytes, clears placeholder, and returns relationship_id' do
      result = doc.replace_image_by_placeholder('{{photo}}', StringIO.new('new-bytes'))

      expect(result[:relationship_id]).to eq('rId5')
      expect(result[:entry_path]).to eq('word/media/image1.png')
      expect(doc.instance_variable_get(:@replace)['word/media/image1.png']).to eq('new-bytes')

      paragraph_text = doc.doc.at_xpath('//w:body/w:p', Docx::Document::XML_NAMESPACES)
                          .xpath('.//w:t', Docx::Document::XML_NAMESPACES).map(&:text).join
      expect(paragraph_text).not_to include('{{photo}}')
    end

    it 'applies cover fit by setting srcRect crop values' do
      doc.replace_image_by_placeholder('{{photo}}', StringIO.new(fake_png(1000, 1000)), fit: :cover, cleanup_placeholder: false)

      src_rect = doc.doc.at_xpath('//pic:blipFill/a:srcRect', Docx::Document::XML_NAMESPACES)
      expect(src_rect).not_to be_nil
      expect(src_rect['t'].to_i).to be > 0
      expect(src_rect['b'].to_i).to be > 0
    end

    it 'keeps placeholder text when cleanup_placeholder is false' do
      doc.replace_image_by_placeholder('{{photo}}', StringIO.new('new-bytes'), cleanup_placeholder: false)

      paragraph_text = doc.doc.at_xpath('//w:body/w:p', Docx::Document::XML_NAMESPACES)
                          .xpath('.//w:t', Docx::Document::XML_NAMESPACES).map(&:text).join
      expect(paragraph_text).to include('{{photo}}')
    end

    it 'raises ImagePlaceholderNotFound when placeholder is missing' do
      expect { doc.replace_image_by_placeholder('{{missing}}', StringIO.new('x')) }
        .to raise_error(Docx::Errors::ImagePlaceholderNotFound)
    end

    it 'raises ImageNotFound when placeholder exists but host has no image' do
      no_image_doc_xml = <<~XML
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p><w:r><w:t>{{photo}}</w:t></w:r></w:p>
          </w:body>
        </w:document>
      XML
      no_image_doc = build_doc(no_image_doc_xml)

      expect { no_image_doc.replace_image_by_placeholder('{{photo}}', StringIO.new('x')) }
        .to raise_error(Docx::Errors::ImageNotFound)
    end
  end

  describe 'in a table cell host (regression)' do
    let(:table_doc_xml) do
      <<~XML
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                    xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
          <w:body>
            <w:tbl>
              <w:tr>
                <w:tc>
                  <w:p><w:r><w:t>{{photo}}</w:t></w:r></w:p>
                  <w:p>
                    <w:r>
                      <w:drawing>
                        <wp:inline>
                          <wp:extent cx="1600" cy="900" />
                          <a:graphic>
                            <a:graphicData>
                              <pic:pic>
                                <pic:blipFill>
                                  <a:blip r:embed="rId5" />
                                  <a:stretch><a:fillRect/></a:stretch>
                                </pic:blipFill>
                                <pic:spPr>
                                  <a:xfrm>
                                    <a:ext cx="1600" cy="900" />
                                  </a:xfrm>
                                </pic:spPr>
                              </pic:pic>
                            </a:graphicData>
                          </a:graphic>
                        </wp:inline>
                      </w:drawing>
                    </w:r>
                  </w:p>
                </w:tc>
              </w:tr>
            </w:tbl>
          </w:body>
        </w:document>
      XML
    end
    let(:doc) { build_doc(table_doc_xml) }

    it 'replaces image when placeholder and image are in different paragraphs of the same cell' do
      allow(zip).to receive(:find_entry).with('word/media/image1.png').and_return(true)

      result = doc.replace_image_by_placeholder('{{photo}}', StringIO.new('cell-bytes'))

      expect(result[:relationship_id]).to eq('rId5')
      expect(doc.instance_variable_get(:@replace)['word/media/image1.png']).to eq('cell-bytes')

      cell_text = doc.doc.xpath('//w:tc//w:t', Docx::Document::XML_NAMESPACES).map(&:text).join
      expect(cell_text).not_to include('{{photo}}')
    end
  end
end
