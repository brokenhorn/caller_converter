import os.path

class FileManager:
    result = []

    def __init__(self, file_path):
        self.file_path = file_path
        self.file_name = os.path.basename(file_path)
        self.new_file_path = None

    def manage(self):

        return self.new_file_path