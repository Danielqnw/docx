require 'docx/containers/table_cell'
require 'docx/containers/container'

module Docx
  module Elements
    module Containers
      class TableColumn
        include Container
        include Elements::Element

        def self.tag
          'w:gridCol'
        end

        def initialize(cells)
          @node = ''
          @properties_tag = ''
          @cells = cells
        end

        # Array of cells contained within row
        def cells
          @cells
        end
        
      end
    end
  end
end
