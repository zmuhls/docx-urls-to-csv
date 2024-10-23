# docx-urls-to-csv
This Python script extracts text and URLs from all .docx files in a directory (excluding temporary files) and writes the file names (via h1) along with the URLs into a CSV file.

## how to run the script
To run the extract_docx_url.py script from the command line, follow these steps:

	1.	Navigate to the directory where the extract_docx_url.py script and the .docx files are located using the cd command:

cd /path/to/directory


	2.	Ensure Python is installed by running:

python3 --version


	3.	Install the necessary python-docx package if you haven’t already:

pip install python-docx


	4.	Run the script using the following command:

python3 extract_docx_url.py

This will execute the script, process all .docx files in the directory, extract URLs, and save them into a CSV file (e.g., output_urls.csv).
