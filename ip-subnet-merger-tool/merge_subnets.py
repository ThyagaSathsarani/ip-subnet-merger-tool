import pandas as pd
import ipaddress
import os
import re
import pdfplumber

def extract_subnets_from_text(text):
    subnet_pattern = r'\b(?:\d{1,3}\.){3}\d{1,3}/\d{1,2}\b'
    return re.findall(subnet_pattern, text)

def read_subnets(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    subnets = []

    if ext in ['.xlsx', '.xls']:
        df = pd.read_excel(file_path)
        subnet_strs = df.iloc[:, 0].dropna().tolist()
    elif ext == '.pdf':
        subnet_strs = []
        try:
            with pdfplumber.open(file_path) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        subnet_strs.extend(extract_subnets_from_text(text))
        except Exception as e:
            print(f"Error reading PDF: {file_path} — {e}")
            subnet_strs = []
    else:
        print(f"Unsupported file type: {file_path}")
        return subnets

    for s in subnet_strs:
        try:
            subnets.append(ipaddress.IPv4Network(s.strip()))
        except ValueError:
            print(f"Skipping invalid subnet: {s}")
    return subnets

def merge_subnets_partial(subnets):
    subnets = sorted(subnets, key=lambda x: int(x.network_address))
    result = []
    temp_group = []

    def try_merge_group(group):
        if len(group) <= 1:
            return group
        collapsed = list(ipaddress.collapse_addresses(group))
        if len(collapsed) == 1:
            merged_block = collapsed[0]
            total_ips = sum(net.num_addresses for net in group)
            if merged_block.num_addresses == total_ips:
                return [merged_block]
        return group

    for subnet in subnets:
        if not temp_group:
            temp_group.append(subnet)
            continue

        test_group = temp_group + [subnet]
        merged_test = try_merge_group(test_group)

        if len(merged_test) == 1:
            temp_group.append(subnet)
        else:
            merged_or_original = try_merge_group(temp_group)
            result.extend(merged_or_original)
            temp_group = [subnet]

    if temp_group:
        merged_or_original = try_merge_group(temp_group)
        result.extend(merged_or_original)

    unique_result = sorted(set(result), key=lambda x: int(x.network_address))
    return unique_result

def write_to_excel(subnets, output_file):
    data = [str(net) for net in subnets]
    df = pd.DataFrame(data, columns=["Subnets"])
    df.to_excel(output_file, index=False)
    print(f"Output written to {output_file}")

def main(*input_files, output_file="output_subnets.xlsx"):
    all_subnets = []
    for file in input_files:
        subnets = read_subnets(file)
        all_subnets.extend(subnets)

    if not all_subnets:
        print("No valid subnets found.")
        return

    merged_subnets = merge_subnets_partial(all_subnets)
    write_to_excel(merged_subnets, output_file)
