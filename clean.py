#!/bin/python3
import pandas as pd

# Define CPT codes and keywords for wound protector and clean closure
wound_protector_cpt = {44204, 44207, 44208, 44205, 44206, 44207, 44208, 44210, 44211, 44212,
                       44140, 44141, 44143, 44144, 44145, 44147, 44160, 45112, 45110, 45119, 45120}
wound_protector_words = {'COLECTOMY', 'COLON RESECTION', 'LOW ANTERIOR BOWEL RESECTION'}

clean_closure_cpt = {44204, 44207, 44208, 44205, 44206, 44207, 44208, 44210, 44211, 44212,
                     44140, 44141, 44143, 44144, 44145, 44147, 44160, 45112, 45110, 45119, 45120,
                     48150, 48152, 48153, 48154, 50840, 50650, 50660, 51590, 51550, 51555, 51565,
                     51570, 51575, 51580, 51585, 51595, 51596, 53210, 53215, 50700, 52341, 52344,
                     44660, 44661}
clean_closure_words = {'COLECTOMY', 'LOW ANTERIOR BOWEL RESECTION', 'COLON RESECTION',
                       'NEPHROURETERECTOMY', 'BOWEL', 'CECOTOMY', 'COLOTOMY', 'ILEOSTOMY', 'CYSTECTOMY'}


def load_data(file_path, sheet_name=0):
    """Load data from an Excel file."""
    return pd.read_excel(file_path, sheet_name=sheet_name)


def filter_procedures(data, cpt_set, keyword_set, debug=False):
    """Filter data based on CPT codes and keywords."""
    def is_applicable(cpt_codes, procedure):
        reasons = []
        for code in str(cpt_codes).split(' , '):
            try:
                if int(code) in cpt_set:
                    reasons.append(f'CPT code: {code}')
            except ValueError:
                continue
        for word in keyword_set:
            if word in procedure.upper().replace(',', '').split():
                reasons.append(f'Keyword: {word}')
        return bool(reasons), reasons

    results = data.apply(lambda row: is_applicable(row['CPT CODES'], row['PRIM PROCEDURE']), axis=1)
    data['APPLICABLE'] = results.apply(lambda x: x[0])
    data['REASONS'] = results.apply(lambda x: x[1])
    return data['APPLICABLE']


def calculate_rate(data, applicable_col, used_col):
    """Calculate the rate of usage."""
    applicable = data[applicable_col].sum()
    used = data[used_col].sum()
    rate = (used / applicable) * 100 if applicable else 0
    return applicable, used, rate


def save_to_excel(data, file_path):
    """Save DataFrame to an Excel file."""
    data.to_excel(file_path, index=False)


def analyze_wound_protector(data, debug=False):
    """Analyze wound protector usage."""
    filter_procedures(data, wound_protector_cpt, wound_protector_words, debug)
    data['USED'] = data['WOUND PROT USED YN'] == 'Yes'
    applicable, used, rate = calculate_rate(data, 'APPLICABLE', 'USED')

    if debug:
        print("Debug Info for Wound Protector:")
        for index, row in data[data['APPLICABLE']].iterrows():
            print(f"PRIM PROCEDURE: {row['PRIM PROCEDURE']}, REASONS: {row['REASONS']}")

    print(f'Wound Protector - Applicable: {applicable}, Used: {used}, Rate: {rate:.2f}%')


def analyze_clean_closure(data, debug=False):
    """Analyze clean closure usage."""
    filter_procedures(data, clean_closure_cpt, clean_closure_words, debug)
    data['USED'] = data.apply(lambda row: row['CLEAN CLOSURE PROC YN'] == 'Yes' if row['APPLICABLE'] else False, axis=1)
    applicable, used, rate = calculate_rate(data, 'APPLICABLE', 'USED')

    if debug:
        print("Debug Info for Clean Closure:")
        for index, row in data[data['APPLICABLE']].iterrows():
            print(f"PRIM PROCEDURE: {row['PRIM PROCEDURE']}, REASONS: {row['REASONS']}")

    print(f'Clean Closure - Applicable: {applicable}, Used: {used}, Rate: {rate:.2f}%')


def extract_min_before_incision(antibiotics_str):
    """Extract MIN B4 INCISION values from the antibiotics string."""
    times = {}
    if pd.isna(antibiotics_str):
        return times
    lines = antibiotics_str.split("\n")
    for line in lines:
        parts = line.split()
        try:
            med_name = parts[0]
            for i, part in enumerate(parts):
                if part == "MIN" and parts[i + 1] == "B4" and parts[i + 2] == "INCISION:":
                    times[med_name] = int(parts[i + 3])
        except (IndexError, ValueError):
            continue
    return times


def calculate_pre_incision_abx_avg(data):
    service_averages = {}

    # Iterate through each row in the data
    for index, row in data.iterrows():
        service = row['SERVICE']
        antibiotics = row['PRE-INCISION ANTIBIOTICS']

        min_b4_incision_times = extract_min_before_incision(antibiotics)

        if service not in service_averages:
            service_averages[service] = {}

        for med, time in min_b4_incision_times.items():
            if med not in service_averages[service]:
                service_averages[service][med] = []
            service_averages[service][med].append(time)

    # Calculate the average timings and include counts
    avg_service_averages = {}
    for service, meds in service_averages.items():
        avg_service_averages[service] = {}
        for med, times in meds.items():
            avg_service_averages[service][med] = {
                'average': sum(times) / len(times) if times else 0,
                'count': len(times)
            }

    print('Service Averages:', avg_service_averages)


def is_colorectal_procedure(procedure, cpt_codes):
    """Check if the procedure is colorectal based on CPT codes and keywords."""
    reasons = []
    for line in str(cpt_codes).split(" , "):
        try:
            if int(line) in wound_protector_cpt:
                reasons.append(f'CPT code: {line}')
        except ValueError:
            continue
    for word in wound_protector_words:
        if word in procedure.upper().replace(',', '').split():
            reasons.append(f'Keyword: {word}')
    return bool(reasons), reasons


def analyze_clean_closure_colorectal(data, debug=False):
    """Analyze clean closure usage specifically for colorectal procedures."""
    results = data.apply(lambda row: is_colorectal_procedure(row['PRIM PROCEDURE'], row['CPT CODES']), axis=1)
    data['IS_COLORECTAL'] = results.apply(lambda x: x[0])
    data['REASONS'] = results.apply(lambda x: x[1])
    data['APPLICABLE'] = data['IS_COLORECTAL']
    data['USED'] = data.apply(lambda row: row['CLEAN CLOSURE PROC YN'] == 'Yes' if row['APPLICABLE'] else False, axis=1)
    applicable, used, rate = calculate_rate(data, 'APPLICABLE', 'USED')

    if debug:
        print("Debug Info for Clean Closure Colorectal:")
        for index, row in data[data['APPLICABLE']].iterrows():
            print(f"PRIM PROCEDURE: {row['PRIM PROCEDURE']}, REASONS: {row['REASONS']}")

    print(f'Clean Closure Colorectal - Applicable: {applicable}, Used: {used}, Rate: {rate:.2f}%')


def main():
    file_path = "/Users/arslanamir/Documents/Work/Quality/Nancy/OR SSI Report All Locations May 2024.xlsx"
    data = load_data(file_path)

    # Analyze wound protector usage
    # analyze_wound_protector(data, debug=True)

    # Analyze clean closure usage
    # analyze_clean_closure(data, debug=True)

    # Analyze clean closure usage specifically for colorectal procedures
    # analyze_clean_closure_colorectal(data, debug=True)

    # Calculate pre and post incision antibiotic averages
    calculate_pre_incision_abx_avg(data)

    # Save results to Excel
    # save_to_excel(data, "/Users/arslanamir/Documents/Work/Quality/practice.xlsx")


if __name__ == "__main__":
    main()
