import compare_excel_to_xml
import convert_excel_to_xml
import search
import time

def menu():

    while (True):
        print("\n\nMenu driven code to convert xls/xlsm to XML and to compare when required ")
        print('\nOptions:')
        print('\t1.convert xls/xlsx into xml')
        print('\t2.compare xls/xlsx with xml')
        print('\t3.Search File')
        print('\t4.Exit\n\n')
        choice = input('Please enter your choice here:')
        print('\n\n')
        if choice == 1 or choice == '1':
            match = convert_excel_to_xml.convert_xls_to_xml()
            if match == 1:
                continue
        elif choice == 2 or choice == '2':
            match = compare_excel_to_xml.check_xml_data()
            #print (match)
            if match == 0:
                print('100 % Match')
            else:
                continue
        elif choice == 3 or choice == '3':
            filename = input('Enter filename to search here(q to quit search):')
            if filename == 'q':
                print('Quitting file search....')
                continue
            else:
                search.file_search(filename)
        elif choice == 4 or choice == '4':
            print('Good Bye!')
            time.sleep(3)
            exit(0)
        else:
            print('Invalid Option')

menu()
