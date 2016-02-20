require 'mail'
require 'roo'
require 'highline/import'

VALID_EMAIL_REGEX = /\A[\w+\-.]+@[a-z\d\-]+(\.[a-z]+)*\.[a-z]+\z/i

#email:foo@bar.com
#file_name:test

details=Hash.new
details["email"]=ARGV[0].split(":")[1]
ans=(VALID_EMAIL_REGEX =~ (details["email"]))
if ans.nil?
  puts "Entered email ID is not valid. Are you using the correct format?"
  puts "Format=> ruby send_emails.rb email:foo@bar.com file_name:/home/foo/Desktop/test"
  exit
end

details["file_name"]=ARGV[1].split(":")[1]

if !(File.exist? details["file_name"]+".xlsx")
  puts "File not found. Are you using the full path to file?"
  puts "Format=> ruby send_emails.rb email:foo@bar.com file_name:/home/foo/Desktop/test"
  exit
end

password = ask("Enter password: ") { |q| q.echo = false }

puts "Filename:" + details["file_name"]

xlsx = Roo::Spreadsheet.open("#{details["file_name"]}.xlsx")
sheet=xlsx.sheet(0)

puts "Document info:"
puts xlsx.info
puts "Sending mails:"
Mail.defaults do
  delivery_method :smtp, { :address    => "smtp.gmail.com",
                          :port       => 587,
                          :user_name  => details["email"],
                          :password   => password,
                          :authentication => :plain,
                          :enable_starttls_auto => true
                        }
end

count=1
sheet.each(full_name: 'Full Name:', email: 'Email:') do |hash|
  
  # => { id: 1, name: 'John Smith' }

  if hash[:full_name]!='Full Name:'
    puts "#{count}. #{hash[:full_name]} (#{hash[:email]})"
    mail = Mail.new do
      from     details["email"]
      to       hash[:email]
      subject  'Your Subject here'
    end

    html_part = Mail::Part.new do
      content_type 'text/html; charset=UTF-8'
      body "<center>
              <h2>Your message here!/h2>
              <br>
            </center>
            <p>Hi #{hash[:full_name]}!</p>
         "
    end

    mail.html_part = html_part
  

    begin
      mail.deliver!
    rescue => e
      puts "Unable to send email because #{e.message}"
      puts "Cannot continue further."
      exit
    end
    count=count+1
  end


end

puts "All mails sent!. Total mails: #{count-1}"