module Html2Docx
  module DocumentObjects
    class Image
      def initialize(document, relation, tmp_path)
        @document = document
        @relation = relation
        @tmp_path = tmp_path

        @media_path = nil
        @image = nil

        check_and_create_media_directory
      end

      def add_image(image_object)
        image = get_image_information(image_object)
        drawing_tag = create_drawing_tag
        inline_tag = create_inline_tag
        doc_pr_tag = create_doc_pr_tag(image)
        graphic_tag = create_graphic_tag(image)
        extent_tag = create_extent_tag(image)
        c_nv_graphic_frame_pr = create_c_nv_graphic_frame_pr(image)

        inline_tag.add_child(extent_tag)
        inline_tag.add_child(doc_pr_tag)
        inline_tag.add_child(c_nv_graphic_frame_pr)
        inline_tag.add_child(graphic_tag)
        drawing_tag.add_child(inline_tag)

        drawing_tag
      end

      private

      def get_image_information(image_object)
        id = @relation.get_latest_image_id + 1
        path = image_object.attr('src')
        name = path.split('/').last
        title = image_object.attr('alt') || "Picture-#{id}"
        height = image_object.attr('height').to_i
        width = image_object.attr('width').to_i

        { name: name,  title: title, path: path, height: height, width: width }
      end

      def create_drawing_tag
        Nokogiri::XML::Node.new('w:drawing', @document)
      end

      def create_inline_tag
        anchor_tag = Nokogiri::XML::Node.new('wp:inline', @document)

        anchor_tag
      end

      def create_doc_pr_tag(image)
        doc_pr_tag = Nokogiri::XML::Node.new('wp:docPr', @document)
        doc_pr_tag['id'] = @relation.get_uniq_image_id
        doc_pr_tag['name'] = image[:name]
        doc_pr_tag['title'] = image[:title]

        doc_pr_tag
      end

      def check_and_create_media_directory
        @media_path = File.join(@tmp_path, 'media')

        Dir.mkdir @media_path unless Dir.exist? @media_path
      end

      def create_graphic_tag(image)
        graphic_tag = Nokogiri::XML::Node.new('a:graphic', @document)

        graphic_data_tag = create_graphic_data_tag(image)
        graphic_tag.add_child(graphic_data_tag)

        graphic_tag
      end

      def create_graphic_data_tag(image)
        graphic_data_tag = Nokogiri::XML::Node.new('a:graphicData', @document)
        graphic_data_tag['uri'] = 'http://schemas.openxmlformats.org/drawingml/2006/picture'

        pic_tag = create_pic_tag(image)
        graphic_data_tag.add_child(pic_tag)

        graphic_data_tag
      end

      def create_pic_tag(image)
        pic_tag = Nokogiri::XML::Node.new('pic:pic', @document)

        nv_pic_pr_tag = create_nv_pic_pr_tag(image)
        pic_tag.add_child(nv_pic_pr_tag)

        blip_fill_tag = create_blip_fill_tag(image)
        pic_tag.add_child(blip_fill_tag)

        sp_pr_tag = create_sp_pr_tag(image)
        pic_tag.add_child(sp_pr_tag)

        pic_tag
      end

      def create_nv_pic_pr_tag(image)
        nv_pic_pr_tag = Nokogiri::XML::Node.new('pic:nvPicPr', @document)

        c_nv_pr_tag = create_c_nv_pr_tag(image)
        nv_pic_pr_tag.add_child(c_nv_pr_tag)

        c_nv_pic_pr = create_c_nv_pic_pr(image)
        nv_pic_pr_tag.add_child(c_nv_pic_pr)

        nv_pic_pr_tag
      end

      def create_c_nv_pr_tag(image)
        c_nv_pr_tag = Nokogiri::XML::Node.new('pic:cNvPr', @document)
        c_nv_pr_tag['id'] = @relation.get_uniq_image_id
        c_nv_pr_tag['name'] = image[:name]
        c_nv_pr_tag['title'] = image[:title]

        c_nv_pr_tag
      end

      def create_c_nv_pic_pr(image)
        c_nv_pic_pr_tag = Nokogiri::XML::Node.new('pic:cNvPicPr', @document)

        c_nv_pic_pr_tag
      end

      def create_blip_fill_tag(image)
        blip_fill_tag = Nokogiri::XML::Node.new('pic:blipFill', @document)

        blip_tag = create_blip_tag(image)
        blip_fill_tag.add_child(blip_tag)

        stretch_tag = create_stretch_tag
        blip_fill_tag.add_child(stretch_tag)

        blip_fill_tag
      end

      def create_blip_tag(image)
        blip_tag = Nokogiri::XML::Node.new('a:blip', @document)
        blip_tag['r:embed'] = @relation.add_image(image, @media_path)

        blip_tag
      end

      def create_stretch_tag
        stretch_tag = Nokogiri::XML::Node.new('a:stretch', @document)

        fill_rect_tag = create_fill_rect_tag
        stretch_tag.add_child(fill_rect_tag)

        stretch_tag
      end

      def create_fill_rect_tag
        Nokogiri::XML::Node.new('a:fillRect', @document)
      end

      def create_sp_pr_tag(image)
        sp_pr_tag = Nokogiri::XML::Node.new('pic:spPr', @document)

        xfrm_tag = create_xfrm_tag(image)
        sp_pr_tag.add_child(xfrm_tag)

        prst_geom_tag = create_prst_geom_tag(image)
        sp_pr_tag.add_child(prst_geom_tag)

        sp_pr_tag
      end

      def create_xfrm_tag(image)
        xfrm_tag = Nokogiri::XML::Node.new('a:xfrm', @document)

        off_tag = create_off_tag(image)
        xfrm_tag.add_child(off_tag)

        ext_tag = create_ext_tag(image)
        xfrm_tag.add_child(ext_tag)

        xfrm_tag
      end

      def create_off_tag(image)
        off_tag = Nokogiri::XML::Node.new('a:off', @document)
        off_tag['x'] = '0'
        off_tag['y'] = '0'

        off_tag
      end

      def create_ext_tag(image)
        ext_tag = Nokogiri::XML::Node.new('a:ext', @document)
        ext_tag['cx'] = image[:width] * 9525
        ext_tag['cy'] = image[:height] * 9525

        ext_tag
      end

      def create_prst_geom_tag(image)
        prst_geom_tag = Nokogiri::XML::Node.new('a:prstGeom', @document)
        prst_geom_tag['prst'] = 'rect'

        av_lst_tag = create_av_lst_tag(image)
        prst_geom_tag.add_child(av_lst_tag)

        prst_geom_tag
      end

      def create_av_lst_tag(image)
        Nokogiri::XML::Node.new('a:avLst', @document)
      end

      def create_extent_tag(image)
        ext_tag = Nokogiri::XML::Node.new('wp:extent', @document)
        ext_tag['cx'] = image[:width] * 9525
        ext_tag['cy'] = image[:height] * 9525

        ext_tag
      end

      def create_c_nv_graphic_frame_pr(image)
        c_nv_graphic_frame_pr_tag = Nokogiri::XML::Node.new('wp:cNvGraphicFramePr', @document)

        graphic_frame_locks_tag = create_graphic_frame_locks_tag(image)
        c_nv_graphic_frame_pr_tag.add_child(graphic_frame_locks_tag)

        c_nv_graphic_frame_pr_tag
      end

      def create_graphic_frame_locks_tag(image)
        graphic_frame_locks_tag = Nokogiri::XML::Node.new('a:graphicFrameLocks', @document)
        graphic_frame_locks_tag['noChangeAspect'] = 1

        graphic_frame_locks_tag
      end
    end
  end
end