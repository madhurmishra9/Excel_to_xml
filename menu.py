"""Interactive menu front-end for the excel_xml package."""

import time

from excel_xml import (
    check_xml_data,
    excel_to_xml,
    file_search,
    remote_to_excel,
    remote_to_xml,
    xml_to_excel,
)


def _convert_excel_to_xml():
    excel_path = input("Path to the Excel file (.xls/.xlsx): ").strip()
    out_dir = input("Folder to write XML files into: ").strip()
    try:
        written = excel_to_xml(excel_path, out_dir)
        print("Wrote %d XML file(s)." % len(written))
    except Exception as exc:
        print("Error: %s" % exc)


def _convert_xml_to_excel():
    xml_dir = input("Folder containing the XML files: ").strip()
    out_path = input("Output Excel path (blank for default): ").strip()
    try:
        out = xml_to_excel(xml_dir, out_path)
        print("Wrote %s" % out)
    except Exception as exc:
        print("Error: %s" % exc)


def _compare():
    excel_path = input("Path to the Excel file: ").strip()
    xml_dir = input("Folder containing the XML files: ").strip()
    try:
        result = check_xml_data(excel_path, xml_dir)
        print("100 % Match" if result == 0 else "Differences found (see above).")
    except Exception as exc:
        print("Error: %s" % exc)


def _fetch(to_excel):
    source = input("URL or cloud URI (s3://, gs://, az://, gdrive://): ").strip()
    kind = input("Kind [auto/rss/xml/html] (blank = auto): ").strip() or "auto"
    try:
        if to_excel:
            out_path = input("Output Excel path (blank for default): ").strip()
            print("Wrote %s" % remote_to_excel(source, out_path, kind))
        else:
            out_dir = input("Output XML folder (blank for default): ").strip()
            written = remote_to_xml(source, out_dir, kind)
            print("Wrote %d XML file(s)." % len(written))
    except Exception as exc:
        print("Error: %s" % exc)


def _search():
    filename = input("File name to search (q to cancel): ").strip()
    if filename.lower() == "q":
        print("Quitting file search....")
        return
    file_search(filename)


def menu():
    actions = {
        "1": _convert_excel_to_xml,
        "2": _convert_xml_to_excel,
        "3": _compare,
        "4": lambda: _fetch(to_excel=True),
        "5": lambda: _fetch(to_excel=False),
        "6": _search,
    }
    while True:
        print("\n\nExcel <-> XML converter")
        print("\nOptions:")
        print("\t1. Convert Excel -> XML")
        print("\t2. Convert XML -> Excel")
        print("\t3. Compare Excel with XML")
        print("\t4. Fetch feed/URL -> Excel")
        print("\t5. Fetch feed/URL -> XML")
        print("\t6. Search for a file")
        print("\t7. Exit\n")
        choice = input("Please enter your choice here: ").strip()
        print("\n")
        if choice == "7":
            print("Good Bye!")
            time.sleep(1)
            break
        action = actions.get(choice)
        if action:
            action()
        else:
            print("Invalid Option")


if __name__ == "__main__":
    menu()
