from datetime import datetime
from pathlib import Path
our_files = Path("/Users/nikpi/Desktop/Files")
for file in our_files.iterdir():
    # Set up key variables for the parent path and the file extensions
    directory = file.parent
    extension = file.suffix

    # Use unpacking by splitting the old name on the '-' character
    old_name = file.stem
    region, report_type, old_date = old_name.split('-')

    # Convert date to datetime and convert to string of desired format
    old_date = datetime.strptime(old_date, "%Y%b%d")
    date = datetime.strftime(old_date, '%Y-%m-%d') 
    
    # Format as DATE - REGION - REPORT TYPE
    new_name = f'{date} - {region} - {report_type}{extension}'
    
    # Rename the file
    file.rename(Path(directory, new_name))