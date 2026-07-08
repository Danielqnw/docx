# frozen_string_literal: true

require 'spec_helper'
require 'docx'
require 'securerandom'
require 'tmpdir'
require 'zip'
require 'stringio'

describe Docx::Merge::MediaImporter do
  def fixtures_path
    'spec/fixtures'
  end

  def rels_ns
    'http://schemas.openxmlformats.org/package/2006/relationships'
  end

  def content_types_ns
    Docx::Document::CONTENT_TYPES_NS
  end

  def image_rel_type
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
  end

  def hyperlink_rel_type
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'
  end

  def build_doc_with_rels_and_media(rels_xml, media_entries = {})
    base = File.join(fixtures_path, 'basic.docx')
    buffer = Zip::OutputStream.write_buffer do |out|
      Zip::File.open(base) do |zf|
        zf.each do |entry|
          next unless entry.file?

          out.put_next_entry(entry.name)
          if entry.name == 'word/_rels/document.xml.rels'
            out.write(rels_xml)
          elsif media_entries.key?(entry.name)
            out.write(media_entries[entry.name])
          else
            out.write(zf.read(entry.name))
          end
        end

        media_entries.each do |path, bytes|
          next if zf.find_entry(path)

          out.put_next_entry(path)
          out.write(bytes)
        end
      end
    end
    Docx::Document.open(StringIO.new(buffer.string))
  end

  def save_to_tempfile(doc)
    path = File.join(Dir.tmpdir, "docx_merge_#{SecureRandom.hex(8)}.docx")
    doc.save(path)
    path
  end

  def rels_xml(path)
    Zip::File.open(path) do |zip|
      entry = zip.glob('word/_rels/document*.xml.rels').first
      zip.read(entry.name)
    end
  end

  def content_types_xml(path)
    Zip::File.open(path) { |zip| zip.read(Docx::Document::CONTENT_TYPES_PATH) }
  end

  def media_rels_xml(image_rid:, image_target:, duplicate_rid: nil, hyperlink_rid: nil, hyperlink_url: 'https://example.com/test')
    duplicate_target = duplicate_rid ? %(<Relationship Id="#{duplicate_rid}" Type="#{image_rel_type}" Target="#{image_target}"/>) : ''
    hyperlink_rel = hyperlink_rid ? %(<Relationship Id="#{hyperlink_rid}" Type="#{hyperlink_rel_type}" Target="#{hyperlink_url}" TargetMode="External"/>) : ''

    <<~XML
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Relationships xmlns="#{rels_ns}">
        <Relationship Id="#{image_rid}" Type="#{image_rel_type}" Target="#{image_target}"/>
        #{duplicate_target}
        #{hyperlink_rel}
      </Relationships>
    XML
  end

  let(:replacement_bytes) { File.binread(File.join(fixtures_path, 'replacement.png')) }
  let(:source_media_path) { 'word/media/source_image.png' }
  let(:source_image_rid) { 'rId20' }
  let(:duplicate_image_rid) { 'rId21' }
  let(:hyperlink_rid) { 'rId22' }

  let(:source) do
    build_doc_with_rels_and_media(
      media_rels_xml(
        image_rid: source_image_rid,
        image_target: 'media/source_image.png',
        duplicate_rid: duplicate_image_rid,
        hyperlink_rid: hyperlink_rid
      ),
      source_media_path => replacement_bytes
    )
  end

  let(:target) { Docx::Document.open(File.join(fixtures_path, 'basic.docx')) }

  describe '#import' do
    it 'imports internal image relationships with new media parts and content types' do
      existing_rids = target.instance_variable_get(:@rels).xpath('//xmlns:Relationship/@Id').map(&:value)
      replace_before = target.instance_variable_get(:@replace).dup

      importer = described_class.new(target, source)
      new_rid = importer.import(source_image_rid)

      expect(new_rid).to match(/\ArId\d+\z/)
      expect(existing_rids).not_to include(new_rid)
      expect(importer.rid_map[source_image_rid]).to eq(new_rid)

      replace = target.instance_variable_get(:@replace)
      new_media_entry = replace.keys.find { |path| path.match?(%r{\Aword/media/image\d+\.png\z}) }
      expect(new_media_entry).not_to be_nil
      expect(replace[new_media_entry]).to eq(replacement_bytes)
      expect(replace_before.keys).not_to include(new_media_entry)

      relation = target.instance_variable_get(:@rels).at_xpath("//xmlns:Relationship[@Id='#{new_rid}']")
      expect(relation['Type']).to eq(image_rel_type)
      expect(relation['Target']).to eq(new_media_entry.sub(%r{\Aword/}, ''))
      expect(relation['TargetMode']).to be_nil

      content_types = Nokogiri::XML(replace[Docx::Document::CONTENT_TYPES_PATH])
      png_default = content_types.at_xpath("//xmlns:Default[@Extension='png']", 'xmlns' => content_types_ns)
      expect(png_default['ContentType']).to eq('image/png')

      temp_path = save_to_tempfile(target)
      saved_rels = Nokogiri::XML(rels_xml(temp_path))
      saved_relation = saved_rels.at_xpath("//xmlns:Relationship[@Id='#{new_rid}']")
      expect(saved_relation).not_to be_nil
      expect(saved_relation['Target']).to eq(new_media_entry.sub(%r{\Aword/}, ''))

      Zip::File.open(temp_path) do |zip|
        expect(zip.find_entry(new_media_entry)).not_to be_nil
        expect(zip.read(new_media_entry)).to eq(replacement_bytes)
      end
    ensure
      File.delete(temp_path) if defined?(temp_path) && temp_path && File.exist?(temp_path)
    end

    it 'imports external hyperlink relationships without adding media' do
      replace_before_count = target.instance_variable_get(:@replace).size

      importer = described_class.new(target, source)
      new_rid = importer.import(hyperlink_rid)

      expect(new_rid).to match(/\ArId\d+\z/)
      expect(new_rid).not_to eq(hyperlink_rid)

      relation = target.instance_variable_get(:@rels).at_xpath("//xmlns:Relationship[@Id='#{new_rid}']")
      expect(relation['Type']).to eq(hyperlink_rel_type)
      expect(relation['Target']).to eq('https://example.com/test')
      expect(relation['TargetMode']).to eq('External')

      replace = target.instance_variable_get(:@replace)
      new_media_entries = replace.keys.grep(%r{\Aword/media/})
      expect(new_media_entries).to be_empty
      expect(replace.size).to eq(replace_before_count)
    end

    it 'deduplicates media when multiple source rIds point to the same target' do
      importer = described_class.new(target, source)

      first_rid = importer.import(source_image_rid)
      second_rid = importer.import(duplicate_image_rid)

      expect(first_rid).to eq(second_rid)

      replace = target.instance_variable_get(:@replace)
      media_entries = replace.keys.grep(%r{\Aword/media/image\d+\.png\z})
      expect(media_entries.size).to eq(1)
    end

    it 'returns the original rId when the source relationship is missing' do
      replace_before = target.instance_variable_get(:@replace).dup
      rels_before = target.instance_variable_get(:@rels).serialize

      importer = described_class.new(target, source)
      result = importer.import('rId999')

      expect(result).to eq('rId999')
      expect(importer.rid_map['rId999']).to eq('rId999')
      expect(target.instance_variable_get(:@replace)).to eq(replace_before)
      expect(target.instance_variable_get(:@rels).serialize).to eq(rels_before)
    end

    it 'reuses cached media bytes without reading source zip during import' do
      importer = described_class.new(target, source)
      allow(source.zip).to receive(:read).and_call_original

      new_rid = importer.import(source_image_rid)

      expect(new_rid).to match(/\ArId\d+\z/)
      expect(source.zip).not_to have_received(:read)
    end
  end
end
