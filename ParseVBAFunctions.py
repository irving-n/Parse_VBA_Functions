import re
file = open(
    r'C:\Users\Irving\Desktop\Programming\Access VBA\Submittals\In Progress\qte_form check for unpaired rst.txt')

data = file.read()

search_queue = ['set rst =', 'rst.close', 'set rst = Nothing']

# init_rst_regex = re.compile(r"(.*)(Set rst = )(.*)$", flags=re.IGNORECASE)
init_lst = [(x.start(), x.end(), x.group(0)) for x in re.finditer('Set rst = db', data)]

# init_it = init_rst_regex.finditer(data)
# init_lst = [(x.start(), x.end(), x.group(0)) for x in init_it]

# close_rst_regex = re.compile(r"(.*)(rst.CLOSE)(.*)$", flags=re.IGNORECASE)
# close_it = close_rst_regex.finditer(data)
# close_lst = [(x.start(), x.end(), x.group(0)) for x in close_it]
close_lst = [(x.start(), x.end(), x.group(0)) for x in re.finditer('rst.CLOSE', data)]

# rst_nothing_regex = re.compile(r"(.*)(Set rst = Nothing)(.*)$", flags=re.IGNORECASE)
# nothing_it = rst_nothing_regex.finditer(data)
# nothing_lst = [(x.start(), x.end(), x.group(0)) for x in nothing_it]
nothing_lst = [(x.start(), x.end(), x.group(0)) for x in re.finditer('Set rst = Nothing', data)]

lined = re.compile(r"\n")
line_it = lined.finditer(data)


progn = re.compile(r"(?m)^((Public |Private )?(Sub |Function )(.+)\((.*)\)(.*)$)")
line_num = 1
output = []
for i, m in enumerate(progn.finditer(data)):
    pos_start = m.start()
    pos_end = m.end()
    text = m.group(0)
    line_char = next(line_it).start()
    
    while pos_start > line_char:
        line_num += 1
        try:
            line_char = next(line_it).start()
        except StopIteration:
            break

    print('[%02d]: Line-%02d | %02d-%02d: %s' % (i, line_num-1, pos_start, pos_end, text))


for search in search_queue:
    instance_count = data.upper().count(search.upper())
    print(f'Instances of the string [{search}] occuring in document: {instance_count}')

file.close()
