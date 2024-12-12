import win32com.client
from datetime import datetime

# Get the current year, month, and day dynamically
current_year = datetime.now().year
current_month = str(datetime.now().month).zfill(2)  # Zero-padding month
current_day = str(datetime.now().day).zfill(2)  # Zero-padding day

# Define dynamic file paths for both PowerPoint presentations
# ppt_file_1 = fr"C:\Users\HCLRDP1\OneDrive - heromotors\MIS Project\MIS\Jaskaran\AR_MIS\{current_year}\{current_month}\{current_day}\Output\AI PPT - Jaskaran overall,top25,collection,dealer OD.pptx"
ppt_file_1 = fr"D:\Neelitech\MIS\python\Truck In Transit\New folder\Truck In Transit Status.pptx"

# Initialize PowerPoint application
ppt_app = win32com.client.Dispatch("PowerPoint.Application")
ppt_app.Visible = True


# Function to open, update links, and save a PowerPoint file
def update_ppt(ppt_file):
    try:
        # Open the presentation
        presentation = ppt_app.Presentations.Open(ppt_file)

        # Update all links in the presentation
        presentation.UpdateLinks()

        # Save and close the presentation
        presentation.Save()
        presentation.Close()

        print(f"Successfully updated: {ppt_file}")
    except Exception as e:
        print(f"Failed to update {ppt_file}: {e}")


# Update both PowerPoint presentations
update_ppt(ppt_file_1)
# update_ppt(ppt_file_2)

# Quit PowerPoint application
ppt_app.Quit()
