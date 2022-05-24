import aspose.words as aw
import glob, codecs

i = 0

docs = []
useful_docs = []
txts = []

# !!! BEFORE PUTTING YOUR .DOCX IN DOCS/ RENAME THEM ACCORDING TO WHAT TICKETS THEY CONTAIN !!!
# !!! BEFORE PUTTING YOUR .DOCX IN DOCS/ RENAME THEM ACCORDING TO WHAT TICKETS THEY CONTAIN !!!
# !!! BEFORE PUTTING YOUR .DOCX IN DOCS/ RENAME THEM ACCORDING TO WHAT TICKETS THEY CONTAIN !!!


# get all .docx from /docs
for doc in glob.glob("docs/*.docx"):
    docs.append(doc)


# !!! BEFORE PUTTING YOUR .DOCX IN DOCS/ RENAME THEM ACCORDING TO WHAT TICKETS THEY CONTAIN !!!
# !!! BEFORE PUTTING YOUR .DOCX IN DOCS/ RENAME THEM ACCORDING TO WHAT TICKETS THEY CONTAIN !!!
# !!! BEFORE PUTTING YOUR .DOCX IN DOCS/ RENAME THEM ACCORDING TO WHAT TICKETS THEY CONTAIN !!!

# sort by filename
docs.sort()


for doc in docs:
    useful_docs.append(aw.Document(doc))

# convert to txts
for doc in useful_docs:
    doc.save(f'tmp/output{i}.txt')
    i += 1

# appending txts-list by every txt
for txt in glob.glob("tmp/*.txt"):
    txts.append(txt)

# removing useless copyright strings
for txt in txts:
    with codecs.open(txt, "r+", "utf_8_sig") as f:
        lines = f.readlines()
        del lines[1]
        del lines[-1]
        f.seek(0)
        f.truncate()
        f.writelines(lines)

# compiling all the outputs into the final file
with codecs.open("final.txt", "w", "utf_8_sig") as outfile:
    for filename in txts:
        with codecs.open(filename, encoding="utf_8_sig") as infile:
            contents = infile.read()
            outfile.write(contents)