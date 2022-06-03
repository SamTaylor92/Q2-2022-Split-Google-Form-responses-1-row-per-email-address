# [Q2 2022] Split Google Form spreadsheet responses based on the amount of email addresses entered into a field
<p align="right"> <a 
href="https://stackoverflow.com/users/18680621/sam-taylor" target="_blank"><img alt="StackOverflow" 
src="https://stackoverflow-badge.vercel.app/?userID=18680621" /></a> <a 
href="https://github.com/SamTaylor92" target="_blank"><img alt="Github" 
src="https://img.shields.io/badge/GitHub-181717.svg?style=for-the-badge&logo=GitHub&logoColor=white" /></a> <a 
href="https://www.linkedin.com/in/samjamest" target="_blank"><img alt="LinkedIn" 
src="https://img.shields.io/badge/LinkedIn-0A66C2.svg?style=for-the-badge&logo=LinkedIn&logoColor=white" /></a> <a 
href="https://signal.group/#CjQKIO50NLkjJmSisbgDD4OhRj5lHG7X-SJTOl-Dn8Fkc4FpEhCYdnCVL1ok4DlVNntY3mGe" target="_blank"><img alt="Signal" src="https://img.shields.io/badge/Signal-3A76F0.svg?style=for-the-badge&logo=Signal&logoColor=white"/></a> <a 
href="mailto:samtaylor92@live.co.uk" target="_blank"><img alt="Email" src="https://img.shields.io/badge/Gmail-D14836?style=for-the-badge&logo=gmail&logoColor=white" /></a>
</p>
<p align="right">
  
![Example_outcome](https://user-images.githubusercontent.com/105542266/171631328-dea24fec-85f9-4053-b764-d875b1bf6499.png)

<h2> Tools</h2>
<p>
<a target="_blank"><img alt="Excel" src="https://img.shields.io/badge/Microsoft%20Excel-217346.svg?style=for-the-badge&logo=Microsoft-Excel&logoColor=white"/></a>
<a target="_blank"><img alt="Google Sheets" src="https://img.shields.io/badge/Google%20Sheets-34A853.svg?style=for-the-badge&logo=Google-Sheets&logoColor=white"/></a>
<a target="_blank"><img alt="Google Forms" src="https://img.shields.io/badge/Google%20Forms-7A1FA2.svg?style=for-the-badge&logo=Google&logoColor=white"/></a>
</p>
  
## Project brief
[**Note:** To protect the client's data & confidentiality, all data has been replaced with dummy data]  

<details open>
<summary> <h3>💼[Problem]</h3> </summary>
  
- The client had trouble with a procedure that required people to enter feedback via a [Google Form](https://docs.google.com/forms/d/1S6O072szAOvC3gC9Y0Xg4-_6mz1vwqweOJRUHXM3gis/viewform?edit_requested=true).<br>
- If there was more than 1 person that needed feedback for a specific incident, then people had to report the incident once per person involved, despite the rest of the information being the same.<br>
- This, of course, wasted time. 

</details>
</details>
  
  
<details open>
<summary> <h3>🎯[Desired outcome]</h3> </summary>

- The client wanted to change the procedure to allow multiple email addresses to be entered in the `person(s) receiving feedback` field of the [Google Form](https://docs.google.com/forms/d/1S6O072szAOvC3gC9Y0Xg4-_6mz1vwqweOJRUHXM3gis/viewform?edit_requested=true).
- They also wanted a formula-based solution to split the [Google Form](https://docs.google.com/forms/d/1S6O072szAOvC3gC9Y0Xg4-_6mz1vwqweOJRUHXM3gis/viewform?edit_requested=true) responses into different rows of the [Google Form (responses)](https://docs.google.com/spreadsheets/d/1m78KrpmYFooY-qyqPZyqhWqVIkb9rdX4-7U85oVYeY0/edit?usp=sharing) spreadsheet.
  - They requested a formula-based solution, so that the solution would be dynamic and also so that they didn't need to run or maintain a script.
- The amount of rows to be generated in the [Google Forms (responses)](https://docs.google.com/spreadsheets/d/1m78KrpmYFooY-qyqPZyqhWqVIkb9rdX4-7U85oVYeY0/edit?usp=sharing) spreadsheet is determined by how many email addresses are entered into the `person(s) receiving feedback` section of the Google Forms.
- The client wanted 1 row to be generated for each additional email address added to the specified question in the [Google Form](https://docs.google.com/forms/d/1S6O072szAOvC3gC9Y0Xg4-_6mz1vwqweOJRUHXM3gis/viewform?edit_requested=true) and all other information duplicated.
  - The standard behaviour for Google Forms is to generate 1 row per Google Forms entry. <br>

`Example:`
>If 3 email addresses are entered into the `person(s) receiving feedback` [Google Forms](https://docs.google.com/forms/d/1S6O072szAOvC3gC9Y0Xg4-_6mz1vwqweOJRUHXM3gis/viewform?edit_requested=true) question, then 3 entries need to be created in the [Google Form responses spreadsheet](https://docs.google.com/spreadsheets/d/1m78KrpmYFooY-qyqPZyqhWqVIkb9rdX4-7U85oVYeY0/edit?usp=sharing) (instead of the usual 1 entry)

</details>
</details>  

<details open>
<summary> <h3>💡[Solution]</h3> </summary>
  
`Description:` <br>Created a formula that used regex features to count how many separators were in the desired column and then used that information to duplicate the information X times, accordingly <br> For best results, use the [Google Drive](https://drive.google.com/drive/folders/1WyRiRz61UrZTVK4pKKVhwpSlE0xBss7D?usp=sharing)  folder and view using [Google Sheets](https://docs.google.com/spreadsheets/d/1m78KrpmYFooY-qyqPZyqhWqVIkb9rdX4-7U85oVYeY0/edit?usp=sharing).<br>
`Formula:` 
>=ARRAYFORMULA(QUERY(SPLIT(FLATTEN('Form Responses 1'!A2:A&"|"&'Form Responses 1'!B2:B&"|"&'Form Responses 1'!C2:C&"|"&'Form Responses 1'!D2:D&"|"&'Form Responses 1'!E2:E&"|"&'Form Responses 1'!F2:F&"|"&'Form Responses 1'!G2:G&"|"&'Form Responses 1'!H2:H&"|"&'Form Responses 1'!I2:I&"|"&IFERROR(trim(SPLIT('Form Responses 1'!C2:C,",",0,0)))),"|",0,0),"Select Col1,Col2,Col10,Col4,Col5,Col6,Col7,Col8,Col9,Col3 where Col10 is not null order by Col1"))
  
`Google Drive:` [(Link)](https://drive.google.com/drive/folders/1WyRiRz61UrZTVK4pKKVhwpSlE0xBss7D?usp=sharing)  
`Repository:` [(Link)](https://github.com/SamTaylor92/-Q2-2022--Split-Google-Form-responses--1-row-per-email-address-)<br> 
`Status:` ✅

</details>
</details>

<br>
This repository is to document the solution to this problem.
</p>

</br></br>
© 2022 GitHub, Inc.
Terms
Privacy
