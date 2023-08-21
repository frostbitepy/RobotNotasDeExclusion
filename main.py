from actions.file_actions import extract_data_from_excel, generate_word_files

def main():
    """Main function to orchestrate the RPA automation process."""
    excel_file_path = "input_data/resumen exclusiones.xlsx"
    template_dir = "note_templates/provalor"
    output_dir = "output_data"

    data_list = extract_data_from_excel(excel_file_path)
    generate_word_files(data_list, template_dir, output_dir)

    print("Automation complete!")

if __name__ == "__main__":
    main()