# frozen_string_literal: true

require 'nokogiri'

module Docx
  module Merge
    class MediaImporter
      RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
      MEDIA_PATH_PATTERN = %r{\Aword/media/}i

      attr_reader :rid_map

      def initialize(target_doc, source_doc)
        @target_doc = target_doc
        @source_doc = source_doc
        @rid_map = {}
        @source_path_to_target_rid = {}
        @media_bytes_cache = {}
        @source_rels_index = build_source_rels_index
        preload_media_bytes
      end

      def import(source_rid)
        return @rid_map[source_rid] if @rid_map.key?(source_rid)

        rel = @source_rels_index[source_rid]
        unless rel
          @rid_map[source_rid] = source_rid
          return source_rid
        end

        if rel[:external]
          new_rid = @target_doc.add_relationship(rel[:type], rel[:target], mode: :external)
          @rid_map[source_rid] = new_rid
          return new_rid
        end

        normalized_path = normalize_entry_path(rel[:target])
        unless media_path?(normalized_path)
          @rid_map[source_rid] = source_rid
          return source_rid
        end

        if @source_path_to_target_rid.key?(normalized_path)
          new_rid = @source_path_to_target_rid[normalized_path]
          @rid_map[source_rid] = new_rid
          return new_rid
        end

        bytes = @media_bytes_cache[normalized_path]
        unless bytes
          warn "MediaImporter: missing cached bytes for #{normalized_path}"
          @rid_map[source_rid] = source_rid
          return source_rid
        end

        ext = File.extname(normalized_path).delete_prefix('.').downcase
        target_media_path = next_media_path(ext)
        @target_doc.add_part(target_media_path, nil, bytes)

        content_type = content_type_for(ext, rel[:type])
        @target_doc.ensure_default_content_type(ext, content_type)

        target_relative = target_media_path.sub(%r{\Aword/}, '')
        new_rid = @target_doc.add_relationship(rel[:type], target_relative)
        @rid_map[source_rid] = new_rid
        @source_path_to_target_rid[normalized_path] = new_rid
        new_rid
      end

      private

      def build_source_rels_index
        source_rels = @source_doc.instance_variable_get(:@rels)
        return {} unless source_rels

        source_rels.xpath('//xmlns:Relationship', 'xmlns' => RELS_NS).each_with_object({}) do |rel_node, index|
          rid = rel_node['Id']
          next unless rid

          index[rid] = {
            target: rel_node['Target'],
            type: rel_node['Type'],
            external: rel_node['TargetMode'] == 'External'
          }
        end
      end

      def preload_media_bytes
        @source_rels_index.each_value do |rel|
          next if rel[:external]

          normalized_path = normalize_entry_path(rel[:target])
          next unless media_path?(normalized_path)
          next if @media_bytes_cache.key?(normalized_path)

          begin
            @media_bytes_cache[normalized_path] = @source_doc.zip.read(normalized_path)
          rescue Errno::ENOENT, Zip::Error => e
            warn "MediaImporter: failed to preload #{normalized_path}: #{e.message}"
          end
        end
      end

      def normalize_entry_path(path)
        @source_doc.send(:normalize_entry_path, path)
      end

      def media_path?(path)
        MEDIA_PATH_PATTERN.match?(path)
      end

      def content_type_for(ext, rel_type)
        Docx::Document::CONTENT_TYPE_MAPPINGS[ext] ||
          content_type_from_rel_type(rel_type) ||
          'application/octet-stream'
      end

      def content_type_from_rel_type(rel_type)
        return unless rel_type&.include?('/image')

        ext = rel_type.split('/').last
        Docx::Document::CONTENT_TYPE_MAPPINGS[ext]
      end

      def next_media_path(ext)
        ext = ext.to_s.downcase
        existing = @target_doc.zip.glob('word/media/*').map(&:name) +
                   @target_doc.instance_variable_get(:@replace).keys
        max_index = existing.filter_map do |name|
          match = name.match(%r{\Aword/media/image(\d+)\.[^/]+\z}i)
          match && match[1].to_i
        end.max || 0
        "word/media/image#{max_index + 1}.#{ext}"
      end
    end
  end
end
