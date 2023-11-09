import sys
import re
from docx import Document

def vtt_to_docx(input_file_path, output_file_name):
    document = Document()

    # Create table structure
    table = document.add_table(rows=1, cols=2)
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Speaker ID'
    hdr_cells[1].text = 'Transcription'

    #pattern_time = re.compile(r"\d{2}:\d{2}\.\d{3} --> \d{2}:\d{2}\.\d{3}")
    pattern_time = re.compile(r"(\d{2}:)?\d{2}:\d{2}\.\d{3} --> (\d{2}:)?\d{2}:\d{2}\.\d{3}")

    pattern_speaker = re.compile(r"\[SPEAKER_\d{2}\]:")

    current_speaker = None
    current_text = ""
    start_timestamp, end_timestamp = None, None
    end_timestamp_temp = None  # Initialize the variable

    def add_row_to_table():
        nonlocal current_speaker, current_text, start_timestamp, end_timestamp
        if current_text:
            current_text.replace("WEBVTT", "")
            cells = table.add_row().cells
            cells[0].text = current_speaker if current_speaker else "Unknown Speaker"
            cells[1].text = current_text.strip() + f" ({start_timestamp} --> {end_timestamp})"
            current_text = ""
            start_timestamp, end_timestamp = None, None

    with open(input_file_path, 'r', encoding='utf-8') as file:
        for line in file:
            line = line.strip()
            if pattern_time.match(line):
                end_timestamp_temp = line.split(" --> ")[1]
                if not start_timestamp:  # If start_timestamp is not already set
                    start_timestamp = line.split(" --> ")[0]
            elif pattern_speaker.match(line):
                speaker = line.split(':')[0] + "]"
                if current_speaker and current_speaker != speaker:
                    if end_timestamp_temp:  # Check if variable is set
                        end_timestamp = end_timestamp_temp
                    add_row_to_table()
                current_speaker = speaker
                current_text += " " + line.split(':')[1].strip()
            elif line:  # for lines without speaker IDs and not a timestamp
                #if current_speaker and current_speaker != "Unknown Speaker":
                #    print(current_speaker)
                #    current_speaker = "Unknown Speaker"
                    
                current_text += " " + line
        if end_timestamp_temp:  # Check if variable is set for the last text
            end_timestamp = end_timestamp_temp  
        add_row_to_table()  # Save the last transcript segment

    document.save(output_file_name)

if __name__ == "__main__":
    input_file_path = sys.argv[1]
    output_file_name = input_file_path.split("\\")[-1].split(".")[0] + ".docx"
    vtt_to_docx(input_file_path, output_file_name)
