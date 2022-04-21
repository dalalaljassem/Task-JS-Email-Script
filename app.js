
// Your code goes here

//converting excel to JSON
const excelToJson = require("convert-excel-to-json");
const result = excelToJson({
  sourceFile: "names.xlsx",
});

console.log(result);
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
  html: `<div class="container" style="color: black;display: table;font-family: Georgia, serif;font-size: 24px;text-align: center;">
  <h1 style="color: rgb(16, 81, 165);font-size: 48px; margin: 20px;"><strong> Certificate of Completion</strong></h1>
  <p style=" margin: 20px;"><strong> This certificate is presented to </strong></p>
  <h2 style="border-bottom: 2px solid black;font-size: 32px;font-style: italic; margin: 20px auto;width: 400px;">${name}</h2>
  <h4 style="margin:20px;">for the completion of</h4>
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