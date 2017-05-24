module Sablon
  module Test
    module Assertions
      def assert_docx_equal(expected_path, actual_path)
        msg = <<-error
The generated document does not match the sample. Please investigate file(s): %s.

If the generated document is correct, the sample needs to be updated:
\t cp #{actual_path} #{expected_path}
    error
        #
        expected_contents = parse_docx(expected_path)
        actual_contents = parse_docx(actual_path)
        #
        mismatch = []
        expected_contents.each do |entry_name, exp_cnt|
          mismatch << entry_name if exp_cnt != actual_contents[entry_name]
        end
        fail format(msg, mismatch.join(' ')) unless mismatch.empty?
      end

      def parse_docx(path)
        contents = {}
        Zip::File.open(path).each do |entry|
          contents[entry.name] = entry.get_input_stream.read
        end
        contents
      end
    end
  end
end
