# notes / manual steps:
#
# 0. load ppt into libre office, and save as odp and export as html; both
#    formats should be placed into a working directory with this structure:
#       working_dir/
#           exported_formats/exported_odp/
#           exported_formats/exported_html/
#           import_ppt_script.rb
#
# 1. Run "ruby import_ppt_script.rb"
#
# 2. Import zip file into D2P2


# setup all the files that will be needed for the zip
csv_filenames = [
             "Answers.csv", 
             "Pages.csv", 
             "PagewiseSkills.csv", 
             "Questions.csv", 
             "QuestionSetQuestions.csv", 
             "QuestionSets.csv", 
             "QuestionwiseSkills.csv", 
             "QuizQuestionSets.csv", 
             "Quizzes.csv", 
             "Sections.csv", 
             "Skills.csv", 
             "Tutors.csv"
]
# create all the files we would have, had we exported a tutor from D2P
csv_filenames.collect{ |f| `touch #{f}`}

# this will be replaced by the import routine, just using this to avoid clashes
tutor_id = 999

# build date string with ruby
date_string = Time.now.utc

# build all the file headers
answers_csv_header = "id,body,correct,question_id,created_at,position,description,image\n"
pages_csv_header = "id,tutor_id,section_id,name,description,content,position,created_at,quiz_id\n"
pagewiseskills_csv_header = "id,page_id,skill_id,created_at,tutor_id\n"
questions_csv_header = "id,body,tutor_id,created_at,feedback,question_type\n"
questionsetquestions_csv_header = "id,question_set_id,question_id,position\n"
questionsets_csv_header = "id,name,body,duration,tutor_id,created_at\n"
questionwiseskills_csv_header = "id,tutor_id,question_id,skill_id,created_at\n"
quizquestionsets_csv_header = "id,quiz_id,question_set_id,position\n"
quizzes_csv_header = "id,name,splash_text,feedback_after_question_set,feedback_after_quiz,question_set_position,tutor_id,created_at\n"
sections_csv_header = "id,name,tutor_id,position,created_at\n"
skills_csv_header = "id,name,skill_type,created_at,tutor_id,notes\n"
tutors_csv_header = "id,user_id,name,tagline,about,media_directory,created_at,navBarColor,backgroundImage,adaptive,logoImage,private,progressive\n"

csv_headers = [
  answers_csv_header,
  pages_csv_header,
  pagewiseskills_csv_header,
  questions_csv_header,
  questionsetquestions_csv_header,
  questionsets_csv_header,
  questionwiseskills_csv_header,
  quizquestionsets_csv_header,
  quizzes_csv_header,
  sections_csv_header,
  skills_csv_header,
  tutors_csv_header
]

#zip headers together with their name for easier enum writing
files_with_headers = Hash[csv_filenames.zip(csv_headers)]

# write all the file headers into the files
files_with_headers.keys.each do |f|
  File.write( f, files_with_headers[f] )
end

# Sections.csv
# defines the structure of the tutor; we create single dummy section
sections_write_string = "1,PlaceHolderSection,#{tutor_id},1,#{date_string}"
open('Sections.csv', 'a') { |f| 
  f << sections_write_string
}


# 1. using exported ODF presentation (.odp)
# copy all pictures to the uploaded images dir in working dir; 
`cp -r exported_formats/exported_odp/Pictures/ uploaded_images/`

# get all the presentation text from the xml file
full_presentation_contents = File.read('exported_formats/exported_odp/content.xml')

# split the pres using 'draw:name...'
split_presentation_contents_array = full_presentation_contents.split('draw:name="page')

# use regex to find which page a picture occurs on
page_num_regex = /Pictures\/.*?"/
matches = []
scanned_array = split_presentation_contents_array[1..-1].each do |a|
  match = a.scan(page_num_regex)
  # remove the trailing quote that the regex grabbed
  matches << match.map!{|m| m.chomp('"')}
end

# 2. using the exported html from libreoffice
# libreoffice outputs each slide as an text{n}.html file; turn each of these into a csv file
text_html_files = Dir[ "exported_formats/exported_html/text*"].sort_by{|s| s[/\d+/].to_i }

# grab content from textN.html file
text_html_files.each_with_index do |filename, index|
  f = File.open(filename, "r")
  f_text = ""
  f.each_line do |l| 
    f_text += l.strip 
  end
  f.close()

  # look for the end of the header in the ppt, and grab until end of body
  # this will pull the main content from the html-formatted slide
  page_content = f_text.scan(/\/h1\>(.+)\<\/body/).to_s
  page_content.gsub!(","," ")
  page_content = page_content[3..-4]

  matches[index].each do |m|
    m.gsub!("Pictures/","")
    img_write_string = "<p><img src='/d2p2/uploaded_images/target/#{m}'></p>"
    page_content << img_write_string
  end

  # for page 0 (the title page), write to Tutors.csv
  if index == 0
    open('Tutors.csv', 'a') { |f|
      page_write_string = "#{tutor_id},20,Content,,#{page_content},#{index + 1},#{date_string}"
      f << page_write_string
    }
  else
    open('Pages.csv', 'a') { |f|
      page_write_string = "#{index},#{tutor_id},1,Content,,#{page_content},#{index},#{date_string},\n"
      f << page_write_string
    }
  end

end

# wrap all the csv files into a zip
p `zip tutor_to_upload.zip *csv uploaded_images/*`
csv_filenames.each do |f|
  File.delete(f)
end
`rm -rf uploaded_images/`
