import re
import openpyxl
import os

def parse_vcf(file_path):
    contacts = []
    name = None
    phone = None

    with open(file_path, 'r', encoding='utf-8') as file:
        for line in file:
            line = line.strip()

            if line.startswith("FN:"):
                name = line[3:]

            elif line.startswith("TEL"):
                match = re.search(r":(.+)", line)
                if match:
                    phone = match.group(1)

            elif line == "END:VCARD":
                if name and phone:
                    contacts.append((name, phone))
                name, phone = None, None

    return contacts

def export_to_excel(contacts, output_file):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Contacts"

    sheet['A1'] = "Name"
    sheet['B1'] = "Mobile Number"

    for i, (name, phone) in enumerate(contacts, start=2):
        sheet[f"A{i}"] = name
        sheet[f"B{i}"] = phone

    workbook.save(output_file)
    print(f"\n✅ Exported {len(contacts)} contacts to:\n{output_file}\n")

def main():
    print("🔼 Please upload your .vcf file in Pydroid 3 first.")
    file_name = input("📄 Enter full VCF file path (e.g. /storage/emulated/0/Download/contacts.vcf): ").strip()

    if not os.path.exists(file_name):
        print(f"❌ File not found: {file_name}")
        return

    contacts = parse_vcf(file_name)
    if not contacts:
        print("⚠️ No contacts found in the VCF file.")
        return

    # Set output file to Downloads folder with same base name
    base_name = os.path.basename(file_name).replace(".vcf", ".xlsx")
    output_file = f"/storage/emulated/0/Download/{base_name}"
    
    export_to_excel(contacts, output_file)

if __name__ == "__main__":
    main()