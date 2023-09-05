# KVS Activity Management Software

## Introduction

The KVS Activity Management Software is a Python-based desktop application developed using the Tkinter library for the Kendriya Vidyalaya Sangathan (KVS). It is designed to streamline the process of managing and organizing school activities, including registrations, result declarations, and information retrieval.

## Features

- **User-Friendly Interface**: The software offers an intuitive and user-friendly interface, making it easy for both teachers and administrators to use.

- **Activity Registration**: Teachers and administrators can register new activities, providing details such as activity name, type, associated teachers, and dates.

- **Result Declaration**: After an activity is completed, the software allows for the declaration of results, including the winner, first runner-up, and second runner-up.

- **Data Storage**: All activity data is stored in an Excel file ("activities.xlsx"), ensuring data integrity and easy access.

- **Information Retrieval**: Users can retrieve information about registered activities and their results, facilitating quick and efficient data retrieval.

## Installation

1. Clone or download the project from the GitHub repository: [GitHub Repository Link](https://github.com/hammadali1805/activity_management).

2. Install the required Python packages:
   ```
   pip install openpyxl
   pip install pywin32
   ```

3. Run the application:
   ```
   python main.py
   ```

## How to Use

### Register a New Activity

1. Click on the "REGISTER" option in the menu.

2. Fill in the required details for the new activity, including name, type, interhouse status, in-charge teacher, associated teachers, and start/end dates.

3. Click the "REGISTER" button to save the activity.

### Declare Activity Results

1. Click on the "DECLARE" option in the menu.

2. Select the activity for which you want to declare results from the dropdown menu.

3. Enter the details for the winner, first runner-up, and second runner-up, including names, houses, and classes.

4. Click the "DECLARE" button to save the results.

### Fetch Activity Information

1. Click on the "FETCH" option in the menu.

2. Select the activity for which you want to retrieve information from the dropdown menu.

3. Click the "FETCH" button to view the details of the selected activity, including its name, type, interhouse status, in-charge teacher, associated teachers, start date, end date, winner, first runner-up, and second runner-up.

## Contributors

- [Hammad Ali](https://github.com/hammadali1805) - Project Lead and Developer

## Acknowledgments

- Special thanks to the Kendriya Vidyalaya Sangathan for the inspiration behind this project.

## Support

If you encounter any issues or have questions about the software, please feel free to [create an issue](https://github.com/hammadali1805/activity_management/issues) on the GitHub repository, and we will be happy to assist you.

Happy managing your school activities with the KVS Activity Management Software!
