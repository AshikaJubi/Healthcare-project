# Exercise-14 Healthcare Automation Project

### Reg No : 212221040020

## AIM: 
To develop an automated system for booking patient appointments and sending email confirmations. This system simplifies appointment scheduling for patients and automates confirmation communication, improving the efficiency of healthcare service.

## Activities Required:
1. Excel Application Scope
2. Read Range
3. For Each Row
4. Assign (for each required patient detail)
5. SMTP Mail Message
6. Log Message


## Procedure:

1. **Open Excel File**: Use the `Excel Application Scope` to open the Excel file containing patient appointment data.

2. **Read Data**: Use the `Read Range` activity within the Excel scope to read all rows of data from the sheet and store it in a DataTable (e.g., `appointmentData`).

3. **Loop Through Appointments**: Use the `For Each Row` activity to iterate through each row in the DataTable, where each row represents a patient’s appointment details.

4. **Extract Patient Information**: Inside the loop, use `Assign` activities to extract and assign values for the patient’s name, email, appointment date, time, and doctor’s name from each row.

5. **Send Confirmation Email**: Use the `SMTP Mail Message` activity to send a confirmation email, filling in the patient's email and customizing the subject and body with extracted appointment details.

6. **Log Sent Email**: After each email is sent, use the `Log Message` activity to record a confirmation that the email was sent to the patient.

7. **Error Handling** (Optional): Wrap the email-sending process in a `Try Catch` block to handle any potential errors, logging any issues that occur.

8. **Close Excel File**: After processing all rows, use the `Close Workbook` activity to close the Excel file.
## Workflow:


![image](https://github.com/user-attachments/assets/51c72088-b467-474f-bf09-3e56ce87d788)


## Output:
![image](https://github.com/user-attachments/assets/f82c50ae-4ed4-48bd-b9ca-a70a42725933)
![image](https://github.com/user-attachments/assets/afe1fd89-4e61-409d-8f32-0c8ba485ef7c)


## Result:
The system automatically sends personalized confirmation emails to patients after they book appointments via Google Forms. Each email is logged, ensuring all appointments are acknowledged promptly without manual intervention.
