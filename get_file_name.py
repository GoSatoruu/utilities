import os

def get_all_files(directory):
    # List all files and directories in the given directory
    files = [f for f in os.listdir(directory) if os.path.isfile(os.path.join(directory, f))]
    return files

def save_to_txt(file_names, output_file):
    # Save the list of file names into a text file with UTF-8 encoding
    with open(output_file, 'w', encoding='utf-8') as file:
        for name in file_names:
            file.write(name + '\n')  # Write each file name on a new line

# Example usage
directory = 'c:/Users/phima/Kenjaku_Research/Pharse 2'  # Replace with your directory path
output_file = 'file_names.txt'  # The name of the text file to save to

file_names = get_all_files(directory)
save_to_txt(file_names, output_file)

print(f"File names have been saved to {output_file}")
