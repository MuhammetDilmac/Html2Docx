module Html2Docx
  module Helpers
    class DocumentHelper
      def self.height_px_to_word(value)
        (value.to_i * 9533.77880184).to_i
      end

      def self.convert_hex_color(value)
        value.upcase.delete('#')
      end

      def self.px_to_indent(value)
        value.to_i.to_i * 15
      end

      def self.width_px_to_word(value)
        (value.to_i * 9405.9375).to_i
      end

      def self.line_height(value)
        (value.to_f * 240).to_i
      end
    end
  end
end