
// Your code goes here

//converting excel to JSON
const excelToJson = require("convert-excel-to-json");
const result = excelToJson({
  sourceFile: "names.xlsx",
});


// initializing my variables (DD/MM/YYYY)
let date = new Date();
let month = date.getMonth();
let year = date.getFullYear();
let day = date.getDate();
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

// Sending my mail using sendGrid
const sgMail = require('@sendgrid/mail')
sgMail.setApiKey(process.env.SENDGRID_API_KEY)
const msg = {
  to: 'dalal.aljassem@gmail.com', // Change to your recipient -> change my email to email variable
  from: 'dalal.aljassem@gmail.com', // Change to your verified sender
  subject: 'Certificate of completion',
  text: 'bla', 
  html: `<div class="container">
  <h1><strong> Certificate of Completion</strong></h1>
  <p><strong> This certificate is presented to </strong></p>
  <h2>${name}</h2>
  <h4>for the secssful completion of</h4>
  <h3 style="margin:20px;">${info}</h3>
  <h4>${day}/${month}/${year}</h4>
  <h5>With a passing grade of ${grade} </h5>
</div>`,
}
 
sgMail
  .send(msg)
  .then(() => { console.log('Email sent')})
  .catch((error) => { console.error(error)})
})