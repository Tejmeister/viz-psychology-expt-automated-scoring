# viz-psychology-expt-automated-scoring

## Background
Psychology students are expected to perform experiments and take survey on various topics to enhance their knowledge. One such survery was done by my friend.
The survey was done on Google Forms and the inputs were collected in an excel sheet. After the inputs were received the student was expected to score the inputs for each participant manually based on certain criteria. Since the number of students was huge it was quite a tedious and error-prone process to score all the participants manually.
This repo scores the participants manually defined by various criteria as defined under metadata folder.

## How to use
1. Install python 3.x
2. Install the libraries using pip
```
pip install -r requirements.txt
```
3. Execute the scoring-main.py with 2 arguments.
	a. input_responses_file_path: File path to the input excel sheet containing responses
	b. output_responses_file_path: File path to the output excel sheet containing scores

```
python scoring-main.py --input_responses_file_path <INPUT_FILE_PATH> --output_responses_file_path <OUTPUT_FILE_PATH>
```