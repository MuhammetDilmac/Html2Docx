module Html2Docx
  module DocumentObjects
    class Heading
      def initialize(document, relation)
        @document = document
        @relation = relation
        @heading  = nil
      end

      def add_heading(heading_object)
        heading_object['class'] = "Heading#{heading_object.name.scan(/[0-9]/).first}"
        heading_object.name = 'p'

        paragraph = Paragraph.new(@document, @relation)
        paragraph.add_paragraph(heading_object)

        @heading = paragraph.render
      end

      def render
        @heading
      end
    end
  end
end