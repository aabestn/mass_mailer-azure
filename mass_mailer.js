var nodemailer = require('nodemailer');
var smtpTransport = require('nodemailer-smtp-transport');
var xlsx = require('xlsx');

var transporter = nodemailer.createTransport(smtpTransport({
    service: 'gmail',
    host: 'smtp.gmail.com',
    auth: {
        user: 'aaryanp2304@gmail.com', // process.env.REACT_APP_EMAIL
        pass: 'kfrd owey upxg lvuf' // process.env.REACT_APP_EMAIL_PASS
    }
}));

var workbook = xlsx.readFile('./LBS.xlsx'); // Replace with the path to your Excel file
var sheetName = workbook.SheetNames[0];
var worksheet = workbook.Sheets[sheetName];

// Extract data from the Excel file
var emailData = xlsx.utils.sheet_to_json(worksheet);

var counter = 0;
var emailsSent = 0;
var totalEmailsToSend = 50; // Set the desired number of emails to send

var i = setInterval(function () {
    if (counter < emailData.length && emailsSent < totalEmailsToSend) {
        sendEmail(emailData[counter]);
        counter++;
        emailsSent++;
    } else {
        clearInterval(i);
    }
}, Math.random() * 5000);

function sendEmail(row) {
    var mailOptions = {
        from: '"Aaryan Panda | IIT-Delhi" ', // Replace with your actual Gmail email address
        to: row.Email, // Use the correct case based on your Excel column header
        subject: 'Regarding request to pursue a research project under You',
        html: `<p>Respected Professor ${row.Name},</p>

        <p>I am Aaryan Panda, a second-year undergraduate student of <b>Biotechnology and Biochemical Engineering</b> at the <b>Indian Institute of Technology, Delhi</b>. I’m writing to apply for a summer research opportunity for may 2024. I feel that such an opportunity would provide me with the ability to apply my knowledge to real-life settings and assist me in gaining an in-depth understanding of these areas, allowing me to attempt to continue with them as my interests.
        </p>

        <p>I've always been interested in business and economics. I have participated in some case study competitions. I am also an active member of the Economics Club and Entrepreneurship Development Club of my University. I've also completed many other courses, including online courses on Machine Learning, investment risk management Project Management, and Marketing Analytics, theories of personality . Besides, I am familiar with Python, SQL, and C++ programming languages, which will allow me to use Data Analytics and Marketing strategies in the business and economics realm.
        </p>
        <p>Your academic effort and publications amazed and fascinated me as I browsed your website. I was hoping for some internship opportunities available for the winter of 2023-24 so that I could broaden my knowledge and gain more experience under your guidance. Due to high travel expenses, I prefer working online, but if funds are provided, I can travel.
    </p>
        <p>I am a dedicated student who is willing to put in every amount of effort that is needed. I'm also willing to learn something ahead of time should you want me to. My POR’s and participation in extracurricular activities have aided in the development of communication and management skills. I can assure you that with my diligence, any project with which I am involved would yield positive results.
        </p>
        <p>I have attached my resume with this email which explains in detail my experiences and achievements. I am willing to oblige to any other requirements you may place: Interviews, Transcripts. Thank you for your time and consideration. I am hoping for a positive response from your side.
    </p>
        Regards <br>
        AARYAN PANDA<br> 


    B. Tech, Biotechnology and Biochemical Engineering, Indian Institute of Technology, Delhi, India<br>


    Contact Number: +91-8595870727<br>


    E-mail address: aaryanp2304@gmail.com, bb1221395@dbeb.iitd.ac.in <br>`,

        attachments: [
            {
                filename: 'Aaryan-CV.pdf', // Replace with the desired file name
                path: './Aaryan-CV.pdf' // Replace with the actual file path
            }
            // Add more attachments if needed
        ]
    };

    // Check if the 'to' field is defined
    if (!mailOptions.to) {
        console.error(`Error: No recipients defined for row: ${JSON.stringify(row)}`);
        return;
    }

    transporter.sendMail(mailOptions, function (error, info) {
        if (error) {
            console.error(`Error sending email to ${row.Email}:`, error);
            // Handle the error according to your needs (e.g., stop the script, log it, etc.)
        } else {
            console.log(`Email sent to ${row.Email}: ${info.response}`);
        }
    });
}