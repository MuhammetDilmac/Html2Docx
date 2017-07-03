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

        paragraph_id     = paragraph_object.attr('id')

        add_paragraph_style paragraph_object
        add_bookmark_start_tag(paragraph_id) if paragraph_id
        add_paragraph_child paragraph_object.children
        add_bookmark_end_tag(paragraph_id) if paragraph_id
      end

      def add_bookmark_start_tag(name)
        bookmark_start_tag = @relation.create_internal_link_start_tag(name, @document)
        @paragraph.add_child(bookmark_start_tag)
      end

      def add_bookmark_end_tag(name)
        bookmark_end_tag = @relation.create_internal_link_end_tag(name, @document)
        @paragraph.add_child(bookmark_end_tag)
      end

      def add_paragraph_style(paragraph_object)
        paragraph_style  = Nokogiri::XML::Node.new('w:pPr',  @document)
        paragraph_styles = []

        style_attribute  = paragraph_object.attr('style')
        style_attributes = style_attribute.split(';') if style_attribute

        paragraph_class  = paragraph_object.attr('class')
        paragraph_class  = paragraph_class.split(' ')&.first if paragraph_class

        if style_attributes
          style_attributes.each do |style|
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
                paragraph_styles.push add_paragraph_line_height(value)
            end
          end
        end

        unless paragraph_class.nil?
          paragraph_styles.push add_paragraph_class(paragraph_class)
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

      def add_paragraph_class(value)
        class_tag = Nokogiri::XML::Node.new('w:pStyle', @document)
        class_tag['w:val'] = value

        class_tag
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

      def add_paragraph_line_height(value)
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
            when 'a'
              href = child.attr('href')
              hyperlink_tag = create_hyperlink_tag(href)
              text_field.add_child(add_link_class(text_style))
              hyperlink_tag.add_child(text_field)
              text_field.add_child add_paragraph_text(child.text)
              hyperlink_tag.add_child text_field
              @paragraph.add_child hyperlink_tag
              next
          end

          paragraph_id = child.attr('id')

          add_bookmark_start_tag(paragraph_id) if paragraph_id
          text_field.add_child add_paragraph_text(child.text)
          add_bookmark_end_tag(paragraph_id) if paragraph_id
          @paragraph.add_child text_field
        end
      end

      def create_text_field
        Nokogiri::XML::Node.new('w:r',  @document)
      end

      def create_text_style
        Nokogiri::XML::Node.new('w:rPr', @document)
      end

      def create_hyperlink_tag(destination)
        hyperlink_tag = Nokogiri::XML::Node.new('w:hyperlink', @document)

        if destination.start_with?('#')
          hyperlink_tag['w:anchor'] = destination.delete('#')
        else
          hyperlink_tag['r:id'] = @relation.create_external_link_id(destination)
        end

        hyperlink_tag
      end

      def add_paragraph_text(value)
        plain_text = Nokogiri::XML::Node.new('w:t', @document)
        plain_text['xml:space'] = 'preserve'
        plain_text.content = value

        plain_text
      end

      def add_link_class(text_style)
        r_style_tag = Nokogiri::XML::Node.new('w:rStyle', @document)
        r_style_tag['w:val'] = 'Hyperlink'
        text_style.add_child(r_style_tag)

        text_style
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