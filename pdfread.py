from pypdf import PdfReader
def extract_data_pdf():
    reader=PdfReader("package.pdf")
    print(reader.pages[0].extract_text().splitlines())
    sku_dict={}
    for num in range(len(reader.pages)):
        extract_list=reader.pages[num].extract_text().splitlines()
        target_sku=extract_list[-3]
        if target_sku not in sku_dict.keys():
            sku_dict[target_sku]=1
        else:
            sku_dict[target_sku]+=1

    return sku_dict


print(extract_data_pdf())
