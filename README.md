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

## Project brief  
<p> <p align='center'> 
<img src="https://user-images.githubusercontent.com/105542266/172706476-41cad84e-7390-483f-820e-efd80c6811b6.gif"/>
</p>
  
> **`Notes:`** 
> - To protect the client's data & confidentiality, all data has been replaced with dummy data
> - For best results, use the [Google Drive](https://drive.google.com/drive/folders/1WyRiRz61UrZTVK4pKKVhwpSlE0xBss7D?usp=sharing) folder and view using [Google Sheets](https://docs.google.com/spreadsheets/d/1m78KrpmYFooY-qyqPZyqhWqVIkb9rdX4-7U85oVYeY0/edit?usp=sharing) (as opposed to the GitHub repository)<br>
  
A small project for a client who wanted to use more of Google Sheet's functionality to save customers time when inputting data into a Google Form.<br>
<br> The aim was: 
- To allow multiple email entries in a Google Forms questionnaire, in the front end.
- In the back end, split the data entries once per email address entered, with the rest of the information duplicated.
  
<h2> Tools</h2>
<p>
<a target="_blank"><img alt="Excel" src="https://img.shields.io/badge/Microsoft%20Excel-217346.svg?style=for-the-badge&logo=Microsoft-Excel&logoColor=white"/></a>
<a target="_blank"><img alt="Google Sheets" src="https://img.shields.io/badge/Google%20Sheets-34A853.svg?style=for-the-badge&logo=Google-Sheets&logoColor=white"/></a>
<a target="_blank"><img alt="Google Forms" src="https://img.shields.io/badge/Google%20Forms-7A1FA2.svg?style=for-the-badge&logo=Google&logoColor=white"/></a>
</p>

<details>
<summary> <h2>Table of contents</h2></summary>	

- [Project brief](#project-brief)
- [Problem](#problem)
- [Desired outcome](#desired-outcome)
  - [Input data](#table-input-data)
  - [Ideal data](#table-output-data)  
- [Solution](#solution)
- [Reference material](#reference-material)
  
</details>

<details open>
<summary> <h3>💼[Problem]</h3> </summary>
  
- The client had trouble with a procedure that required people to enter feedback via a [Google Form](https://docs.google.com/forms/d/1S6O072szAOvC3gC9Y0Xg4-_6mz1vwqweOJRUHXM3gis/viewform?edit_requested=true).<br>
- If there was more than 1 person that needed feedback for a specific incident, then people had to report the incident once per person involved, despite the rest of the information being the same. This, of course, wasted time. <br>
<p align='right'><a href="#-tools" target="_blank">⬆</a></p>	
  
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

![Example_outcome](https://user-images.githubusercontent.com/105542266/171631328-dea24fec-85f9-4053-b764-d875b1bf6499.png)   
  
<p align='right'><a href="#-tools" target="_blank">⬆</a></p>	
  
<details>
<summary> <h4>📊[Table: Input data]</h4> </summary>

| **Timestamp**     | **Person giving feedback**                  | **Person receiving feedback**                                                  | **[Question A]** | **[Question B]** | **[Question C]** | **[Question D]** | **[Question E]** | **[Question F]** |
|-------------------|---------------------------------------------|--------------------------------------------------------------------------------|------------------|------------------|------------------|------------------|------------------|------------------|
| 6/2/2022 14:08:01 | commaspace@commaspace.com                   | commaspace@.com, commaspace2@.com                                              | Answer A         | Option 1         | Option 2         | Option 1         | Text             | Text             |
| 6/2/2022 14:08:32 | spacecomma@spacecomma.com                   | spacecomma@.com ,spacecomma2@.com                                              | Text             | Option 2         | Option 2         | Option 2         | Text             | Text             |
| 6/2/2022 14:08:54 | spacecomma@spacecomma.com                   | spacecommaspace@.com , spacecommaspace2@.com                                   | Text             | Option 2         | Option 2         | Option 5         | Text             | Text             |
| 6/2/2022 14:09:27 | 3emails@hotmail.com                         | email1@hotmail.com,email2@hotmail.com,email3@hotmail.com                       | Text             | Option 2         | Option 2         | Option 9         | Text             | Text             |
| 6/2/2022 14:10:02 | 3-emails-different-comma-errors@hotmail.com | email1spacecomma@hotmail.com ,email2commaspace@hotmail.com, email3@hotmail.com | Text             | Option 2         | Option 1         | Option 7         | Text             | Text             |

<p align='right'><a href="#-tools" target="_blank">⬆</a></p>	  
  
</details>

<details>
<summary> <h4>📊[Table: Output data]</h4> </summary>  

| **Timestamp**             | **Person giving feedback**                  | **Person receiving feedback** | **[Question A]** | **[Question B]** | **[Question C]** | **[Question D]** | **[Question E]** | **[Question F]** | **Original email entry**                                                       |
|---------------------------|---------------------------------------------|-------------------------------|------------------|------------------|------------------|------------------|------------------|------------------|--------------------------------------------------------------------------------|
| June 2, 2022  [14:08:01]  | commaspace@commaspace.com                   | commaspace@.com               | Answer A         | Option 1         | Option 2         | Option 1         | Text             | Text             | commaspace@.com, commaspace2@.com                                              |
| June 2, 2022  [14:08:01]  | commaspace@commaspace.com                   | commaspace2@.com              | Answer A         | Option 1         | Option 2         | Option 1         | Text             | Text             | commaspace@.com, commaspace2@.com                                              |
| June 2, 2022  [14:08:32]  | spacecomma@spacecomma.com                   | spacecomma@.com               | Text             | Option 2         | Option 2         | Option 2         | Text             | Text             | spacecomma@.com ,spacecomma2@.com                                              |
| June 2, 2022  [14:08:32]  | spacecomma@spacecomma.com                   | spacecomma2@.com              | Text             | Option 2         | Option 2         | Option 2         | Text             | Text             | spacecomma@.com ,spacecomma2@.com                                              |
| June 2, 2022  [14:08:54]  | spacecomma@spacecomma.com                   | spacecommaspace@.com          | Text             | Option 2         | Option 2         | Option 5         | Text             | Text             | spacecommaspace@.com , spacecommaspace2@.com                                   |
| June 2, 2022  [14:08:54]  | spacecomma@spacecomma.com                   | spacecommaspace2@.com         | Text             | Option 2         | Option 2         | Option 5         | Text             | Text             | spacecommaspace@.com , spacecommaspace2@.com                                   |
| June 02, 2022  [14:09:27] | 3emails@hotmail.com                         | email1@hotmail.com            | Text             | Option 2         | Option 2         | Option 9         | Text             | Text             | email1@hotmail.com,email2@hotmail.com,email3@hotmail.com                       |
| June 02, 2022  [14:09:27] | 3emails@hotmail.com                         | email2@hotmail.com            | Text             | Option 2         | Option 2         | Option 9         | Text             | Text             | email1@hotmail.com,email2@hotmail.com,email3@hotmail.com                       |
| June 2, 2022  [14:09:27]  | 3emails@hotmail.com                         | email3@hotmail.com            | Text             | Option 2         | Option 2         | Option 9         | Text             | Text             | email1@hotmail.com,email2@hotmail.com,email3@hotmail.com                       |
| June 2, 2022  [14:10:02]  | 3-emails-different-comma-errors@hotmail.com | email1spacecomma@hotmail.com  | Text             | Option 2         | Option 1         | Option 7         | Text             | Text             | email1spacecomma@hotmail.com ,email2commaspace@hotmail.com, email3@hotmail.com |
| June 2, 2022  [14:10:02]  | 3-emails-different-comma-errors@hotmail.com | email2commaspace@hotmail.com  | Text             | Option 2         | Option 1         | Option 7         | Text             | Text             | email1spacecomma@hotmail.com ,email2commaspace@hotmail.com, email3@hotmail.com |
| June 2, 2022  [14:10:02]  | 3-emails-different-comma-errors@hotmail.com | email3@hotmail.com            | Text             | Option 2         | Option 1         | Option 7         | Text             | Text             | email1spacecomma@hotmail.com ,email2commaspace@hotmail.com, email3@hotmail.com |

<p align='right'><a href="#-tools" target="_blank">⬆</a></p>	
  
</details>
  
  
</details>
</details>  

<details open>
<summary> <h3>💡[Solution]</h3> </summary>

For best results, use the [Google Drive](https://drive.google.com/drive/folders/1WyRiRz61UrZTVK4pKKVhwpSlE0xBss7D?usp=sharing) folder and view using [Google Sheets](https://docs.google.com/spreadsheets/d/1m78KrpmYFooY-qyqPZyqhWqVIkb9rdX4-7U85oVYeY0/edit?usp=sharing).<br>  
`Google Drive:` [(Link)](https://drive.google.com/drive/folders/1WyRiRz61UrZTVK4pKKVhwpSlE0xBss7D?usp=sharing)  
`Repository:` [(Link)](https://github.com/SamTaylor92/-Q2-2022--Split-Google-Form-responses--1-row-per-email-address-)<br> 
`Formula:` 
```
=ARRAYFORMULA(QUERY(SPLIT(FLATTEN('Form Responses 1'!A2:A&"|"&'Form Responses 1'!B2:B&"|"&'Form Responses 1'!C2:C&"|"&'Form Responses 1'!D2:D&"|"&'Form Responses 1'!E2:E&"|"&'Form Responses 1'!F2:F&"|"&'Form Responses 1'!G2:G&"|"&'Form Responses 1'!H2:H&"|"&'Form Responses 1'!I2:I&"|"&IFERROR(trim(SPLIT('Form Responses 1'!C2:C,",",0,0)))),"|",0,0),"Select Col1,Col2,Col10,Col4,Col5,Col6,Col7,Col8,Col9,Col3 where Col10 is not null order by Col1"))
```    
<p align='right'><a href="#-tools" target="_blank">⬆</a></p>	

<details open> 
  
<summary> <code> Explanation: </code> </summary>
  
> - Using `FLATTEN` allows us to pull all rows together into one single column
> - We then add a divider `|` to each column
>   - This will help to separate the data back into it's original columns later, to undo the effect of `FLATTEN`
> - We can then use a nested `SPLIT` on the desired column and denote a delimiter
> - In our scenario, we want our spreadsheet entries to be split into multiple columns if there are multiple email addresses (separated by a comma) entered into the `Person(s) receiving feedback` column (Column C)
> - `TRIM` before our nested `SPLIT` helps to reduce the effects of input errors 
>   - As we trim any white spaces from around the delimiter (in our example, a comma `,`), this allows users to input the email addresses with spaces, without causing an error. 
> - The `QUERY` helps:
>  - To filter our results from the nested `SPLIT`
>  - To select the columns and column order that we would like  
> - The `ARRAYFORMULA` ensures that the formula is dynamic and works across all rows, which is useful when the sheet will be continuously populated
<p align='right'><a href="#-tools" target="_blank">⬆</a></p>	

</details>  
</details>
</details>

<details open>
<summary> <h3>📚[Reference material]</h3> </summary>
  
- [x] [Google Documentation](https://support.google.com/docs/answer/10307761?hl=en)
- [x] [Google Support Forums](https://support.google.com/docs/thread/46674663/can-anyone-help-me-to-split-the-date-stamp-and-text-in-the-string-%E2%80%9Capproval-history-status-app?hl=en)    
- [x] [Table Convert](https://tableconvert.com/)
<p align='right'><a href="#-tools" target="_blank">⬆</a></p>	
  
</details>
</details>

</p>

</br></br>
© 2022 GitHub, Inc.
Terms
Privacy
