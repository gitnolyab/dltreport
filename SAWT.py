import streamlit as st
import os
import pandas as pd
import process_dat_file


# ok na ni bqy, complete running Sep 18

def sawt_user_input_path():
    st.write('SAWT - DAT FILE PRINTING')
    # Get the user's input directory and filename
    user_input_directory = st.text_input("ENTER EXACT DIRECTORY  example - D:\\BIR\\SAWT")
    user_input_filename = st.text_input("ENTER CORRECT FILENAME  example -  EXCEL_SAMPLE_DATA_D.xlsx")
    # user_input_SheetName = st.text_input("Enter Sheet Name")

    if user_input_directory and user_input_filename:
        # Combine the user's input directory and filename to create the custom file path
        custom_file_path = os.path.join(user_input_directory, user_input_filename)

        # Check if the custom file path exists and is a file
        if os.path.exists(custom_file_path) and os.path.isfile(custom_file_path):
            # Display the custom file path
            st.write(f"Inputted File Path: {custom_file_path}")

            # Read the Excel content into a DataFrame
            df = pd.read_excel(custom_file_path, engine='openpyxl')

            # Display the DataFrame
            st.write("DataFrame from Custom File:")
            st.dataframe(df, hide_index=True)

            button_clicked = st.button('Process DAT File')
            if button_clicked:
                process_dat_file.process_dat(custom_file_path, user_input_directory)
                # st.success('Processing DAT file...')

        else:
            # Display the "File not found" message if the file is not found
            st.write("File not found,,, Directory or Filename is not correct")
            st.write("Please try to input the Directory and Filename correctly...")


if __name__ == "__main__":
    sawt_user_input_path()
