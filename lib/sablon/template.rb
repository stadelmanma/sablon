module Sablon
  class Template
    def initialize(path)
      @path = path
      # reading all entries into a hash to for quick access while processing
      @contents = {}
      Zip::File.open(@path).each do |entry|
        content = entry.get_input_stream.read
        @contents[entry.name] = if entry.name =~ /\.(?:xml|rels)$/
                                  Nokogiri::XML(content)
                                else
                                  content
                                end
      end
    end

    # Same as +render_to_string+ but writes the processed template to +output_path+.
    def render_to_file(output_path, context, properties = {})
      File.open(output_path, 'wb') do |f|
        f.write render_to_string(context, properties)
      end
    end

    # Process the template. The +context+ hash will be available in the template.
    def render_to_string(context, properties = {})
      render(context, properties).string
    end

    private

    def render(context, properties = {})
      # initialize environment
      env = Sablon::Environment.new(self, context)
      env.relationships.initialize_rids(@contents)
      env.footnotes.initialize_footnotes(@contents['word/footnotes.xml'])
      env.bookmarks.initialize_bookmark_ids(@contents['word/document.xml'])
      # process files
      process(%r{word/document.xml}, env, properties)
      process(%r{word/(?:header|footer)\d*\.xml}, env)
      process(%r{word/footnotes\.xml}, env)
      process(%r{word/endnotes\.xml}, env)
      process(%r{word/numbering.xml}, env)
      process(/\[Content_Types\].xml/, env)
      #
      Zip::OutputStream.write_buffer(StringIO.new) do |out|
        generate_output_file(out, env)
      end
    end

    # IMPORTANT: Open Office does not ignore whitespace around tags.
    # We need to render the xml without indent and whitespace.
    def generate_output_file(zip_out, env)
      # update relationships
      env.relationships.output_new_rids(@contents)
      # output updated zip and add images
      @contents.each do |entry_name, xml_node|
        zip_out.put_next_entry(entry_name)
        if entry_name =~ /\.(?:xml|rels)$/
          zip_out.write(xml_node.to_xml(indent: 0, save_with: 0))
        else
          zip_out.write(xml_node)
        end
      end
      env.images.add_images_to_zip!(zip_out)
    end

    def get_processor(entry_name)
      if entry_name == 'word/document.xml'
        Processor::Document
      elsif entry_name =~ %r{word/(?:header|footer)\d*\.xml}
        Processor::Document
      elsif entry_name =~ %r{word/endnotes\.xml}
        Processor::Document
      elsif entry_name =~ %r{word/footnotes\.xml}
        Processor::Footnotes
      elsif entry_name == 'word/numbering.xml'
        Processor::Numbering
      elsif entry_name == '[Content_Types].xml'
        Processor::ContentType
      end
    end

    def process(entry_pattern, env, *args)
      @contents.each do |entry_name, content|
        next unless entry_name =~ entry_pattern
        env.current_entry = entry_name
        processor = get_processor(entry_name)
        processor.process(content, env, *args)
      end
    end
  end
end
