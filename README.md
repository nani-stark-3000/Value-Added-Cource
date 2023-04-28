# Value Added Course

This is a Flask web application that allows course instructors to create, register students, and mark attendance. The application also provides a feedback section for students to give feedback on the course content.

## Requirements

* Python 3.x
* Flask
* xlsxwriter
* openpyxl
* pandas
* boto3
* opencv-python
* AWS Account

## Installation

1. Clone this repository or download the ZIP file.
2. Install the requirements using `pip install -r requirements.txt`.
3. Create a file named `secret.py` and add your AWS Access Key, AWS Secret Access Key and Region Name as shown below.
   ```python
   access_key = 'YOUR_AWS_ACCESS_KEY'
   secret_access_key = 'YOUR_AWS_SECRET_ACCESS_KEY'
   region_name = 'YOUR_AWS_REGION_NAME'
   ```
4. Start the Flask application by running the following command:
   ```
   python app.py
   ```
5. Navigate to `http://localhost:5000` in your web browser.

## Usage

1. Home page will have two options i.e create course and course list. If the user wants to create a new course, click on the create course button and enter the course name, year, and branch. This will create an excel sheet with attendance and feedback tabs and the excel sheet will be uploaded to the S3 bucket.
2. Once the course is created, the course will be displayed in the course list. Click on the course name to access the attendance sheet.
3. In the attendance sheet, the instructor can register students by entering their name and registration number. The instructor can also mark the attendance of the students by clicking on the respective day's button.
4. In the feedback sheet, the instructor can view the feedback given by the students on various parameters.
5. The instructor can add a new day by clicking on the 'Add Day' button on the attendance sheet.
6. The instructor can download the attendance sheet by clicking on the 'Download' button on the attendance sheet.

## User-Interface

1. Login Page

![login](https://github.com/nani-stark-3000/Value-Added-Cource/blob/2e8e7b0a07e30b4415078741f15508d1c6d94bd2/screenshots/Login.png)

2. Registration Page

![register](https://github.com/nani-stark-3000/Value-Added-Cource/blob/2e8e7b0a07e30b4415078741f15508d1c6d94bd2/screenshots/Registration.png)

3. Student Profile Page

![sp](https://github.com/nani-stark-3000/Value-Added-Cource/blob/2e8e7b0a07e30b4415078741f15508d1c6d94bd2/screenshots/Student%20Profile.png)

4. Teacher Profile Page

![tp](https://github.com/nani-stark-3000/Value-Added-Cource/blob/2e8e7b0a07e30b4415078741f15508d1c6d94bd2/screenshots/Teacher%20Profile.png)

5. Course Registration Page

![cr](https://github.com/nani-stark-3000/Value-Added-Cource/blob/2e8e7b0a07e30b4415078741f15508d1c6d94bd2/screenshots/s_course.png)

6. Courses list View Page

![cl](https://github.com/nani-stark-3000/Value-Added-Cource/blob/2e8e7b0a07e30b4415078741f15508d1c6d94bd2/screenshots/Course%20List.png)

7. Course Creation Page 

![cc](https://github.com/nani-stark-3000/Value-Added-Cource/blob/2e8e7b0a07e30b4415078741f15508d1c6d94bd2/screenshots/course_creation.png)

8. Course view for Teachers

![cvt](https://github.com/nani-stark-3000/Value-Added-Cource/blob/2e8e7b0a07e30b4415078741f15508d1c6d94bd2/screenshots/t_course.png)
