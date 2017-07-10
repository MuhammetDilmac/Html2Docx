module Html2Docx
  class Relation
    def initialize(options = {})
      @relation_file = nil
      @relation = nil
      @relations = []
      @last_relation_id = 1
      @internal_links = {}
      @external_links = {}
      @images = {}
      @unique_image_id = 0

      if options[:main_relation]
        @relation_file = File.join(options.fetch(:temp), 'word', '_rels', 'document2.xml.rels')
        @relation = File.open(@relation_file) { |f| Nokogiri::XML(f) }
        @last_relation_id = @relation.css('Relationship').last.attr('Id').to_i
      else
        @relation_file = File.join(options.fetch(:temp), 'word', '_rels', options.fetch(:file_name))
        @relation = create_relation_file
      end

      @relations = @relation.at_css('Relationships')

      @relation.at_css('Relationship').children.each do |children|
        children.remove
      end
    end

    def create_relation_file
      document = Nokogiri::XML::Document.new
      document.encoding = 'UTF-8'
      relations_tag = Nokogiri::XML::Node.new('Relationships', document)
      relations_tag['xmlns'] = 'http://schemas.openxmlformats.org/package/2006/relationships'
      document.add_child relations_tag
      document
    end

    def create_internal_link_start_tag(name, document)
      bookmark_start_tag = Nokogiri::XML::Node.new('w:bookmarkStart', document)
      bookmark_start_tag['w:id'] = create_internal_link_id(name)
      bookmark_start_tag['w:name'] = name

      bookmark_start_tag
    end

    def create_internal_link_end_tag(name, document)
      bookmark_end_tag = Nokogiri::XML::Node.new('w:bookmarkEnd', document)
      id, value = find_internal_link_id(name)
      bookmark_end_tag['w:id'] = value

      bookmark_end_tag
    end

    def create_internal_link_id(name)
      id = find_internal_link_id(name)
      if id
        id
      else
        id = get_latest_internal_link_id + 1
        @internal_links[id] = name
      end
    end

    def get_latest_internal_link_id
      @internal_links.keys.max || 0
    end

    def find_internal_link_id(name)
      @internal_links.find{ |key, value| value == name }
    end

    def create_external_link_id(destination)
      id, value = find_external_link_id(destination)

      if id
        id
      else
        id = get_latest_external_link_id.delete('rId').to_i + 1
        @external_links["rId#{id}"] = destination
        "elId#{id}"
      end
    end

    def find_external_link_id(destination)
      @external_links.find { |key, value| value == destination }
    end

    def get_latest_external_link_id
      @external_links.keys.max || 'rId0'
    end

    def get_uniq_image_id
      @unique_image_id = @unique_image_id + 1
    end

    def add_image(image, media_path)
      real_path    = image[:path]
      image_name   = image[:path].split('/').last
      current_path = File.join(media_path, image_name)

      if real_path.start_with? 'http'
        request    = Typhoeus::Request.new(path)
        image_file = File.open(current_path, 'wb+')

        request.on_headers do |response|
          if response.code != 200
            raise "Image not found! Image Path: #{real_path}"
          end
        end

        request.on_body do |data|
          image_file.write(data)
        end

        request.on_complete do |response|
          image_file.close
        end

        request.run
      else
        if File.exist? real_path
          FileUtils.cp real_path, current_path
        else
          raise "Image not found! Image Path: #{real_path}"
        end
      end

      relation_image_path = File.join('/', 'media', image_name)

      add_image_relation(relation_image_path)
    end

    def add_image_relation(relation_image_path)
      image_id = "iId#{get_uniq_image_id}"

      @images[image_id] = {
        type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
        target: relation_image_path
      }

      image_id
    end

    def get_latest_image_id
      @images.keys.max.to_i
    end

    def render
      @external_links.each do |key, value|
        external_link_relation = Nokogiri::XML::Node.new('Relationship', @relation)
        external_link_relation['Type'] = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'
        external_link_relation['Target'] = value
        external_link_relation['TargetMode'] = 'External'
        external_link_relation['Id'] = key

        @relation.root.add_child(external_link_relation)
      end

      @images.each do |key, value|
        image_relation = Nokogiri::XML::Node.new('Relationship', @relation)
        image_relation['Type'] = value[:type]
        image_relation['Target'] = value[:target]
        image_relation['Id'] = key

        @relation.root.add_child(image_relation)
      end

      File.open(@relation_file, 'w') { |f| f.write(Helpers::NokogiriHelper.to_xml(@relation)) }
    end
  end
end