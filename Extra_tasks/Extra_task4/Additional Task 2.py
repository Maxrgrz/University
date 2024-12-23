import re
import urllib.request
import csv

site = urllib.request.urlopen('https://msk.spravker.ru/avtoservisy-avtotehcentry/')
content = site.read().decode()

pattern = r"(?:-link\">)(?P<Name>[^<]+)"\
    r"(?:[^o]*[^l]*.*\n *(?P<Address>[^\n]+))"\
    r"(?:\s*.*>\s*.*>\s*.*>(?:\s*<d[^>]*>(?:\s*.*\s*.*>(?P<Phone_Number>[^<]+))?.*>\s*</dl>)"\
    r"(?:\s*<.*>(?:\s*<.*\s*<.*>(?P<WorkHours>[^<]+))?</dd>)?)?"

matches = re.findall(pattern, content)


file = open('Site_GET_data.csv', 'w', newline='', encoding='utf8') 
writer = csv.writer(file)
writer.writerow(['Name', ' Address', ' Phone number', ' Work hours'])
writer.writerows(matches)
file.close()
