
// Your code goes here

//converting excel to JSON
const excelToJson = require("convert-excel-to-json");
const result = excelToJson({
  sourceFile: "names.xlsx",
});

// initializing my variables (DD/MM/YYYY)
let grade = "A";
let date = new Date();
let month = date.getMonth();
let year = date.getFullYear();
let day = date.getDate();

// Sending my mail 
const sgMail = require('@sendgrid/mail')
sgMail.setApiKey(process.env.SENDGRID_API_KEY)
const msg = {
  to: 'dalal.aljassem@gmail.com', // Change to your recipient -> change to email variable
  from: 'dalal.aljassem@gmail.com', // Change to your verified sender
  subject: 'Certificate of completion',
  text: 'and easy to do anywhere, even with Node.js',
  html: '<strong>and easy to do anywhere, even with Node.js</strong>', // -> change to html page
}
sgMail
  .send(msg)
  .then(() => {
    console.log('Email sent')
  })
  .catch((error) => {
    console.error(error)
  })

// looping through my JSON 
result.Sheet1.forEach(function(obj){
    // console.log(obj.A); //name
    // console.log(obj.B); //email
    // console.log(obj.C); //html css js
    // console.log(obj.D); //grade
    let name = obj.A;
    let email = obj.B;
    let info = obj.C;
    let grade = obj.D;
})