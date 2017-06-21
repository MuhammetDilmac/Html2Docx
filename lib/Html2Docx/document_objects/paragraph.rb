module Html2Docx
  module DocumentObjects
    class Paragraph
      def initialize(document, relation)
        @document  = document
        @relation  = relation
        @paragraph = nil
      end

      def add_paragraph(paragraph_object)
        create_paragraph(paragraph_object)
      end

      def create_paragraph(paragraph_object)
        @paragraph = Nokogiri::XML::Node.new('w:p', @document)

        paragraph_style = paragraph_object.attr('style')
        add_paragraph_style paragraph_style if paragraph_style

        add_paragraph_child paragraph_object.children
      end

      def add_paragraph_style(style_attribute)
        paragraph_style = Nokogiri::XML::Node.new('w:pPr',  @document)
        paragraph_styles = []

        styles = style_attribute.split(';')

        styles.each do |style|
          style = style.strip
          attribute, value = style.scan(/(.+):\s?(.+);?/).flatten

          case attribute
            when 'text-indent'
              paragraph_styles.push add_paragraph_indent(value)
            when 'text-align'
              paragraph_styles.push add_paragraph_alignment(value)
            when 'background-color'
              paragraph_styles.push add_paragraph_background_color(value)
            when 'line-height'
              paragraph_styles.push add_line_height(value)
          end
        end

        paragraph_styles.each do |style|
          paragraph_style.add_child(style)
        end

        @paragraph.add_child(paragraph_style)
      end

      def add_paragraph_indent(value)
        indent_tag = Nokogiri::XML::Node.new('w:ind',  @document)
        indent_tag['w:firstLine'] = Helpers::DocumentHelper.px_to_indent(value)

        indent_tag
      end

      def add_paragraph_alignment(value)
        align_tag = Nokogiri::XML::Node.new('w:jc', @document)
        value = value.downcase
        value = 'both' if value == 'justify'
        align_tag['w:val'] = value

        align_tag
      end

      def add_paragraph_background_color(value)
        background_tag = Nokogiri::XML::Node.new('w:shd', @document)
        background_tag['w:val'] = 'clear'
        background_tag['w:color'] = 'auto'
        background_tag['w:fill'] = Helpers::DocumentHelper.convert_hex_color(value)

        background_tag
      end

      def add_line_height(value)
        line_height_tag = Nokogiri::XML::Node.new('w:spacing', @document)
        line_height_tag['w:line'] = Helpers::DocumentHelper.line_height(value)

        line_height_tag
      end

      def add_paragraph_child(children)
        children.each do |child|
          text_field = create_text_field
          text_style = create_text_style

          case child.name
            when 'strong'
              text_field.add_child add_strong_text(text_style)
            when 'i'
              text_field.add_child add_italic_text(text_style)
            when 'font'
              color = child.attr('color')
              text_field.add_child add_font_color(text_style, color) unless color.nil?
            when 'u'
              text_field.add_child add_underline_text(text_style)
            when 's'
              text_field.add_child add_stroke_text(text_style)
          end

          text_field.add_child add_paragraph_text(child.text)
          @paragraph.add_child text_field
        end
      end

      def create_text_field
        Nokogiri::XML::Node.new('w:r',  @document)
      end

      def create_text_style
        Nokogiri::XML::Node.new('w:rPr', @document)
      end

      def add_paragraph_text(value)
        plain_text = Nokogiri::XML::Node.new('w:t', @document)
        plain_text['xml:space'] = 'preserve'
        plain_text.content = value

        plain_text
      end

      def add_strong_text(text_style)
        strong_text = Nokogiri::XML::Node.new('w:b', @document)
        text_style.add_child(strong_text)

        text_style
      end

      def add_italic_text(text_style)
        italic_text = Nokogiri::XML::Node.new('w:i', @document)
        text_style.add_child(italic_text)

        text_style
      end

      def add_font_color(text_style, color)
        color_text = Nokogiri::XML::Node.new('w:color', @document)
        color_text['w:val'] = Helpers::DocumentHelper.convert_hex_color(color)
        text_style.add_child(color_text)

        text_style
      end

      def add_underline_text(text_style)
        underline_text = Nokogiri::XML::Node.new('w:u', @document)
        underline_text['w:val'] = 'single'
        text_style.add_child(underline_text)

        text_style
      end

      def add_stroke_text(text_style)
        stroke_text = Nokogiri::XML::Node.new('w:dstrike ', @document)
        stroke_text['w:val'] = true
        text_style.add_child(stroke_text)

        text_style
      end

      def render
        @paragraph
      end
    end
  end
end