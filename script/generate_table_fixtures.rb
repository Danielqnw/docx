#!/usr/bin/env ruby
# frozen_string_literal: true

require_relative '../spec/support/table_fixture_builder'

output_dir = File.expand_path('../spec/fixtures/tables', __dir__)
TableFixtureBuilder.write_all_fixtures(output_dir)

puts "Generated #{TableFixtureBuilder::FIXTURES.size} fixtures in #{output_dir}:"
TableFixtureBuilder::FIXTURES.each_key do |filename|
  puts "  - #{filename}"
end
