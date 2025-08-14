from kato_sorter import KATOFileSorter

if __name__ == '__main__':
    sorter = KATOFileSorter(input_dir='.')
    sorter.delete_kato_subfolders()
    sorter.process_files()
    sorter.save_kato_files()