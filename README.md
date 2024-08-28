# Data-Labeling
Excel Data Processing and Labeling using Word2Vec
Overview
This project processes and labels data from multiple Excel files using natural language processing techniques, such as Word2Vec for semantic similarity. The main tasks performed by this script include data cleaning, correcting spelling errors, identifying materials, and labeling based on pre-defined keywords using Word2Vec similarity measures.

Features
Data Loading: Loads data from multiple Excel files.
Spelling Correction: Corrects spelling based on a custom dictionary.
Material Identification: Identifies materials mentioned in product descriptions using keywords.
Word2Vec Model Training: Trains a Word2Vec model on product descriptions to capture semantic meaning.
Labeling Based on Similarity: Uses cosine similarity to label text based on its semantic similarity to known keywords.
Physical Property Extraction: Extracts specific physical properties (like thickness, width, etc.) using regular expressions.
Save Labeled Data: Saves the processed and labeled data back to Excel files.
Prerequisites
Ensure you have Python 3.x installed, along with the following packages:

bash
Sao chép mã
pip install pandas openpyxl gensim scikit-learn numpy
Setup
Place Excel Files: Store the input Excel files in the paths specified in the script (import_file_paths and keywords_file_path).
Configure Keywords and Columns: Update the keywords_dict in the script to match the correct sheet names and column names in your Excel files.
How to Run
Train the Word2Vec Model and Process Data:

Run the script from the command line or any Python IDE:

bash
Sao chép mã
python your_script_name.py
The script will:

Load data from the specified Excel files.
Train a Word2Vec model on product descriptions.
Process the data by correcting spelling, identifying materials, and labeling products.
Save the processed data back to Excel files with the suffix _Labeled_Word2Vec_Re.
Functions and Modules
load_data(file_path, sheet_name=None): Loads data from an Excel file.
correct_spelling_with_dict(product): Corrects spelling errors using a predefined dictionary.
train_word2vec(sentences): Trains a Word2Vec model on the provided sentences.
vectorize_text(text, model): Converts text into a vector representation using a trained Word2Vec model.
label_using_word2vec(text, keywords, model, output_column, supply_chain=None): Labels text based on Word2Vec similarity and keyword presence.
identify_materials(text, keywords_df): Identifies materials from text using keyword matching.
label_material(row, material_data, model): Labels materials based on product text.
determine_material_type(description): Determines the type of material based on keywords in the description.
check_special_indicators(text): Checks for special yarn count indicators in the text.
extract_physical_properties(row): Extracts physical properties from text using regex patterns.
detect_new_percentage(text): Detects new percentages in the text.
process_and_label_dataframe(df, model, keywords_dict): Processes the DataFrame and applies all labeling and extraction functions.
save_to_excel(file_path, df, keywords_dict): Saves the processed DataFrame to an Excel file.
Logging
The script uses Python's logging module to log information and errors. Adjust the logging level as needed.

Notes
Ensure that the input Excel files are properly formatted, and the correct sheet names and column names are specified.
The script can handle exceptions, such as missing files or data, and will log appropriate error messages.
Customize the keyword dictionary (keywords_dict) as per your requirements to improve accuracy.
Contributing
Feel free to open issues or pull requests for any bugs or enhancements.

License
This project is licensed under the MIT License - see the LICENSE file for details.
