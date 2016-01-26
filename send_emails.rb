require 'mail'
require 'roo'
puts "enter file name:"
file_name= gets.chomp()
puts file_name

xlsx = Roo::Spreadsheet.open("./#{file_name}.xlsx")
sheet=xlsx.sheet(0)

puts xlsx.info
Mail.defaults do
  delivery_method :smtp, { :address    => "smtp.gmail.com",
                          :port       => 587,
                          #:domain     => 'gmail.com',
                          :user_name  => ENV['EMAIL_ID'],
                          :password   => ENV['GMAIL_PASSWORD'],
                          :authentication => :plain,
                          :enable_starttls_auto => true
                        }
end



sheet.each(full_name: 'Full Name:', email: 'Email:') do |hash|
  puts hash.inspect
  if hash[:full_name]!='Full Name:'
  mail = Mail.new do
    from     'user_name@gmail.com'
    to       hash[:email]
    subject  'Workshop on basics of GIT/GITHUB'
  #  body     'Hey there!'
    #add_file :filename => 'somefile.png', :content => File.read('/somefile.png')
  end

  html_part = Mail::Part.new do
    content_type 'text/html; charset=UTF-8'
    body "<center>
              <h2>Basics of GIT/GITHUB</h2>
              <br>
              <img src=\"http://s19.postimg.org/6rc7kco8j/130712_git_github_topdenota1_compressed_1.jpg\"/>
            </center>
            <p>Hi #{hash[:full_name]}, thank you for registering!</p>
            <p>To get an idea what is Git or GitHub you can go through the following links:</p>
            <ul>
              <li>
                <a href=\"https://www.quora.com/What-is-git-and-why-should-I-use-it\">What is git and why should we use it?<a>
              </li>
              <li>
                <a href=\"https://guides.github.com/activities/hello-world/\">Beginners guide for Github.</a>
              </li>
              <li>
                <a href=\"http://stackoverflow.com/questions/13321556/difference-between-git-and-github\">Difference-between-git-and-github</a>
              </li>
              <li>
                <a href=\"https://try.github.io/levels/1/challenges/1\">Practice GIT online</a>
              </li>
            </ul>
            <br>
            <strong>Make an account on <a href=\"www.github.com\">Github</a> before coming to the session.</strong>
            <p>
            <strong>What you need to bring?</strong><br>If you have a laptop and your own internet connection then bring, otherwise not necessary.</p>
            <br>
            <p>
              <strong>Time: 2:00pm onwards</strong>
              <br>
              <strong>Date: 29th January,2016</strong>
              <br>
              <strong>Venue: BVS Auditorium</strong>

              <center>
                <i>
                  <strong>Be on time. Keep Coding :)</strong>
                </i>
              </center>
             </p>"
  end

  mail.html_part = html_part

  end

end
