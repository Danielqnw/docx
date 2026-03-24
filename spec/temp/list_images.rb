#!/usr/bin/env ruby
# frozen_string_literal: true

require 'docx'

docx_path = ARGV[0]

if docx_path.nil? || docx_path.strip.empty?
  warn "Usage: ruby spec/temp/list_images.rb <docx_path>"
  exit 1
end

doc = Docx::Document.open(docx_path)
images = doc.images

if images.empty?
  puts "No images found in: #{docx_path}"
  exit 0
end

puts "Images in #{docx_path}:"
puts '-' * 72
puts format('%-6s %-12s %s', 'INDEX', 'RID', 'ENTRY_PATH')
puts '-' * 72

images.each_with_index do |(rid, entry_path), index|
  puts format('%-6d %-12s %s', index + 1, rid, entry_path)
end

puts '-' * 72
puts "Total: #{images.size}"
