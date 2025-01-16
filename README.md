# Appointment Scheduling App

## Overview

This project is a tool developed during my internship for a bathroom remodeling company. The app efficiently assigns in-home appointments to sales representatives based on the following factors:

1. **Availability**: Ensures representatives are scheduled only during their working hours.
2. **Product Scope**: Matches appointments to sales reps with relevant expertise.
3. **Proximity**: Minimizes drive time between locations.

## Features

- **Automated Scheduling**: Assigns appointments based on rep availability, expertise, and location.
- **Route Optimization**: Reduces travel time by assigning reps to appointments closest to their current location.
- **User-Friendly Interface**: Easy-to-use interface suitable for employees without technical expertise.

## Technology Stack

- **Backend**: Python 
- **Frontend**: Tkinter (GUI for user interaction)
- **Data Management**: Excel files for storing and managing appointment data

## Installation

This project requires Python 3.6 or higher.

Follow these steps to install and run the app:

1. Clone the repository:
   ```bash
   git clone https://github.com/Blue-Eml/appointment-scheduling-app.git
   ``` 

2. Navigate to the project directory:
   ``` 
   cd appointment-scheduling-app
   ``` 

3. Install the required dependencies:
   ``` 
   pip install -r requirements.txt
   ``` 

4. Run the application:
   ``` 
   python main.py  
   ``` 

## API Key Requirement

This project uses an API for route optimization and location services. When you first run the application, a popup will appear prompting you to enter a valid API key.

To Generate an API Key: 

1.  Visit Google Cloud Console (or approproate platform)
2. Create an API key: 
3. Enable Necessary APIs:  
   - **Maps JavaScript API**
   - **Geocoding API** 
   - **Directions API** 
   - **Distance Matrix API** 

4. Copy Key and Paste into pop-up when prompted 

## App Usage 

1. **Launch the Application:**
   Run the `main.py` script:
   ``` 
   python main.py  
   ``` 

2. **API Key Prompt:**

Enter a valid API key into the popup when prompted.

3. **Upload Appointment data**

Upload the Excel file containing appointment information. 

4. **Enter Employee Data**


5. **Run the Application:**
The app will schedule appointments based on the provided data.

6. **Modify Sales Representatives:**

After assigning appointments between time windows, a popup will ask if you'd like to modify the sales reps, including all the current sale reps. This allows you to change sale reps availability between morning, noon, and evening appointments. 

To Add Sales Reps:

   - Type `add` and select OK.
   - Enter data for one sales rep and press OK.
   - Repeat for additional sales reps, as needed.

To Remove Sales Reps:

   - Type `remove` and select OK.
   - Enter the names of sales reps to remove, separated by commas `(e.g., John, Jane, Sally)` and press OK.

Select No if no modifications are needed, and proceed to assign the next time window.

7. **View Results**

Once the last appointment time window is assigned, results will appear in the terminal, and the final schedule will be saved as an Excel file named after the date:

   ``` 
   appointments_results_YYYY-MM-DD.xlsx 
   ``` 