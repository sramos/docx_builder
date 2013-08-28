module Word
  class WordDocument < Package
    attr_accessor :main_doc

    def initialize(filename)
      super(filename)

      main_doc_part = get_relationship_target(WORD_MAIN_DOCUMENT_TYPE)
      raise PackageError.new("Word document package '#{@filename}' has no main document part") if main_doc_part.nil?
      @main_doc = MainDocument.new(self, main_doc_part)
    end

    def self.blank_document(options={})
      base_document = options.delete(:base_document)
      base_document ||= File.join(File.dirname(__FILE__), 'content', 'blank.docx')
      doc = WordDocument.new(base_document)
      doc.filename = nil
      doc
    end

    def self.from_data(data)
      file = Tempfile.new('OfficeWordDocument')
      file.binmode
      file.write(data)
      file.close
      begin
        doc = WordDocument.new(file.path)
        doc.filename = nil
        return doc
      ensure
        file.delete
      end
    end

    def add_heading(text)
      p = @main_doc.add_paragraph
      p.add_style("Heading1")
      p.add_text_run(text)
      p
    end

    def add_sub_heading(text)
      p = @main_doc.add_paragraph
      p.add_style("Heading2")
      p.add_text_run(text)
      p
    end

    # Add a paragraph to this document
    # Text can be:
    #   A single string such as "text" - inserts that text
    #   A hash such as {:content => "text", :style => "style"} - inserts that text with that character style
    #   An array such as ["text1", "text2"] - inserts each piece of text as a separate run
    #   An array of hashes such as [{:content => "text1", :style => "style1"}, {:content => "text2", :style => "style2"}]
    #      - inserts each piece of text as a separate run with the specified character style
    # NOTE: You may mix strings and hashes in a single array
    # NOTE: styles supplied in the text parameter are character styles, not paragraph styles
    # Available options:
    #   :style - Paragraph style to use
    def add_paragraph(text, options={})
      p = @main_doc.add_paragraph
      style = options.delete(:style)
      if style
        p.add_style(style)
      end
      if text.is_a? Hash
        content = text.delete(:content)
        p.add_text_run(content, text)
      elsif text.is_a? Array
        text.each do |run|
          if run.is_a? Hash
            content = run.delete(:content)
            p.add_text_run(content, run)
          else
            p.add_text_run(run)
          end
        end
      else
        p.add_text_run(text)
      end
      p
    end

    def add_image(image, options={}) # image must be an Magick::Image or ImageList
      p = @main_doc.add_paragraph
      style = options.delete(:style)
      if style
        p.add_style(style)
      end
      p.add_run_with_fragment(create_image_run_fragment(image))
      p
    end

    # keys of hash are column headings, each value an array of column data
    # Available options:
    #   :table_style - Table style to use (defaults to LightGrid)
    #   :column_widths - Array of column widths in twips (1 inch = 1440 twips)
    #   :column_styles - Array of paragraph styles to use per column
    #   :table_properties - A Word::TableProperties object (overrides other options)
    #   :skip_header - Don't output the header if set to true
    def add_table(hash, options={})
      @main_doc.add_table(create_table_fragment(hash, options))
    end

    def plain_text
      @main_doc.plain_text
    end

    # The type of 'replacement' determines what replaces the source text:
    #   Image  - an image (Magick::Image or Magick::ImageList)
    #   Hash   - a table, keys being column headings, and each value an array of column data
    #   Array  - a sequence of these replacement types all of which will be inserted
    #   String - simple text replacement
    def replace_all(source_text, replacement)
      case
      # For simple cases we just replace runs to try and keep formatting/layout of source
      when replacement.is_a?(String)
        @main_doc.replace_all_with_text(source_text, replacement)
      when (replacement.is_a?(Magick::Image) or replacement.is_a?(Magick::ImageList))
        runs = @main_doc.replace_all_with_empty_runs(source_text)
        runs.each { |r| r.replace_with_run_fragment(create_image_run_fragment(replacement)) }
      else
        runs = @main_doc.replace_all_with_empty_runs(source_text)
        runs.each { |r| r.replace_with_body_fragments(create_body_fragments(replacement)) }
      end
    end

    def create_body_fragments(item, options={})
      case
      when (item.is_a?(Magick::Image) or item.is_a?(Magick::ImageList))
        [ "<w:p>#{create_image_run_fragment(item)}</w:p>" ]
      when item.is_a?(Hash)
        [ create_table_fragment(item, options) ]
      when item.is_a?(Array)
        create_multiple_fragments(item)
      else
        [ create_paragraph_fragment(item.nil? ? "" : item.to_s, options) ]
      end
    end

    def create_image_run_fragment(image)
      prefix = ["", @main_doc.part.path_components, "media", "image"].flatten.join('/')
      identifier = unused_part_identifier(prefix)
      extension = "#{image.format}".downcase

      part = add_part("#{prefix}#{identifier}.#{extension}", StringIO.new(image.to_blob), image.mime_type)
      relationship_id = @main_doc.part.add_relationship(part, IMAGE_RELATIONSHIP_TYPE)

      Run.create_image_fragment(identifier, image.columns, image.rows, relationship_id)
    end

    # column_widths option, if supplied, is an array of measurements in twips (1 inch = 1440 twips)
    def create_table_fragment(hash, options={})
      c_count = hash.size
      return "" if c_count == 0

      table_properties = options.delete(:table_properties)
      if table_properties
        table_style = table_properties.table_style
        column_widths = table_properties.column_widths
        column_styles = table_properties.column_styles
      else
        table_style = options.delete(:table_style)
        column_widths = options.delete(:column_widths)
        column_styles = options.delete(:column_styles)
        table_properties = TableProperties.new(table_style, column_widths, column_styles)
      end

      skip_header = options.delete(:skip_header)
      fragment = "<w:tbl>#{table_properties}"

      if column_widths
        fragment << "<w:tblGrid>"
        column_widths.each do |column_width|
          fragment << "<w:gridCol w:w=\"#{column_width}\"/>"
        end
        fragment << "</w:tblGrid>"
      end

      unless skip_header
        fragment <<  "<w:tr>"
        hash.keys.each do |header|
          encoded_header = Nokogiri::XML::Document.new.encode_special_chars(header.to_s)
          fragment << "<w:tc><w:p><w:r><w:t>#{encoded_header}</w:t></w:r></w:p></w:tc>"
        end
        fragment << "</w:tr>"
      end

      r_count = hash.values.inject(0) { |max, value| [max, value.is_a?(Array) ? value.length : (value.nil? ? 0 : 1)].max }
      0.upto(r_count - 1).each do |i|
        fragment << "<w:tr>"
        hash.values.each_with_index do |v, j|
          table_cell = create_table_cell_fragment(v, i,
            :width => column_widths ? column_widths[j] : nil,
            :style => column_styles ? column_styles[j] : nil)
          fragment << table_cell
        end
        fragment << "</w:tr>"
      end

      fragment << "</w:tbl>"
      fragment
    end

    def create_table_cell_fragment(values, index, options={})
      item = case
      when (!values.is_a?(Array))
        index != 0 || values.nil? ? "" : values
      when index < values.length
        values[index]
      else
        ""
      end

      width = options.delete(:width)
      xml = create_body_fragments(item, options).join
      # Word validation rules seem to require a w:p immediately before a /w:tc
      xml << "<w:p/>" unless xml.end_with?("<w:p/>") or xml.end_with?("</w:p>")
      fragment = "<w:tc>"
      if width
        fragment << "<w:tcPr><w:tcW w:type=\"dxa\" w:w=\"#{width}\"/></w:tcPr>"
      end
      fragment << xml
      fragment << "</w:tc>"
      fragment
    end

    def create_multiple_fragments(array)
      array.map { |item| create_body_fragments(item) }.flatten
    end

    def create_paragraph_fragment(text, options={})
      style = options.delete(:style)
      fragment = "<w:p>"
      if style
        fragment << "<w:pPr><w:pStyle w:val=\"#{style}\"/></w:pPr>"
      end
      fragment << "<w:r><w:t>#{Nokogiri::XML::Document.new.encode_special_chars(text)}</w:t></w:r></w:p>"
      fragment
    end

    def debug_dump
      super
      @main_doc.debug_dump
      #Logger.debug_dump_xml("Word Main Document", @main_doc.part.xml)
    end
  end
end