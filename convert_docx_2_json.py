import docx2txt
import sys
import re, json
FILE_NAME = {
    "input": "3.docx",
    "output": "3.json",
}
KEY_SEARCH = {
    "start": "Câu ",
    "answer": ["A\.", "B\.", "C\.", "D\."]
}
##################################
#MAIN
##################################
def main(argv):
    contentFile = read_file_docx(FILE_NAME["input"])
    result = {
        "questions": []
    }
    i = 1
    while(True):
        r = get_json_quesiton_by_number(i, contentFile)
        i = i + 1
        if r == None:
            break
        else:
            result["questions"].append(r)
    write_json(FILE_NAME["output"], result)

##################################
#FUNCTION
##################################
def get_string_by_regex(regexStr, contentFile):
    result = re.findall(re.compile(regexStr, re.MULTILINE), contentFile)
    try:
        return result[0]
    except IndexError:
        return False

def read_file_docx(fileName):
    return docx2txt.process(fileName)

def find_content_by_key_start_n_end(keyS, keyE, content):
    try:
        #print("find_content_by_key_start_n_end: " + keyS + " -> " + keyE)
        start = re.search(keyS, content).start()
        end = 0
        if keyE=="":
            end = len(content)
        else:
            end = re.search(keyE, content).start()

        #print("find_content_by_key_start_n_end: " + keyS + ":" + str(start) + " -> " + keyE + ":" + str(end))
        #print(content[start:end])
        return content[start:end]
    except:
        return None

def find_question_by_number(num, content):
    s1 = KEY_SEARCH["start"]+ str(num) + ":"
    e1 = KEY_SEARCH["start"]+ str(num + 1) + ":"
    entireQuestion = find_content_by_key_start_n_end(s1, e1, content)

    if entireQuestion == None:
        print("[NOT FOUND] " + s1 + " -> " + e1)
        return None, None

    e2 = KEY_SEARCH["answer"][0]
    question = find_content_by_key_start_n_end(s1, e2, entireQuestion)

    answers = []
    s3 = e3 = ""
    for k in range(len(KEY_SEARCH["answer"])):
        if k + 1 < len(KEY_SEARCH["answer"]):
            s3 = KEY_SEARCH["answer"][k]
            e3 = KEY_SEARCH["answer"][k + 1]
        else:
            s3 = KEY_SEARCH["answer"][k]
            e3 = ""
        ele = find_content_by_key_start_n_end(s3, e3, entireQuestion)
        
        ele = ele.strip(s3) #remove A. B. C. D.
        ele = re.sub(r'\t',"", ele) #remove \t
        ele = re.sub(r'\n',"", ele) #remove \n
        answers.append(ele)

    question = question.strip(s1)
    question = re.sub(r'\n',"", question)
    question = re.sub(r'\t',"", question)
    return question, answers

def get_json_quesiton_by_number(num, content):
    question, answers = find_question_by_number(num, content)
    if question==None:
        return None

    result = {
        "question" : question,
        "answers"  : answers
    }
    return result

def write_json(pathFile, data, isMinify = False):
    with open(pathFile, 'w', encoding="utf-8") as outfile:
        if isMinify:
            json.dump(data, outfile, separators=(',',':')) #dump file json format minify
        else:
            json.dump(data, outfile, sort_keys = True, indent = 4, ensure_ascii = False)

###############################
#RUN MAIN
###############################
main(sys.argv)