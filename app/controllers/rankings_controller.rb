class RankingsController < ApplicationController
    require 'rubyXL'

    def new
        if request.post?
            workbook = RubyXL::Parser.parse(params[:file].tempfile.path)
            worksheet = workbook[0]
            our_website = worksheet[1][2].value
            google_websites = worksheet[0][3..-1].map {|cell| cell.value}
            
            worksheet.each_with_index do |row, row_index|
                next if row_index == 0 # skip first row
                keyword = worksheet[row_index][1].value
                google_websites.each_with_index do |website, website_index|
                    # SCRAPPPPPP
                    worksheet.add_cell(row_index, website_index + 3, (website_index + 3).to_s)
                end
            end

            file_path = Rails.root.join('public', 'Google Keywords Ranking.xlsx')
            workbook.write(file_path)
            
            respond_to :js
        end
    end

end