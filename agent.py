from data_canvas import UniversalDataToPPT

def run_agent(input_file):
    output_file = input_file.replace(".csv", "_ai.pptx")

    converter = UniversalDataToPPT()
    converter.convert_file(input_file, output_file)

    return output_file
