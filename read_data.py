import os
import glob
import csv
from docx import Document

def extract_star_paragraphs(paragraphs):
    """提取所有以 * 开头的段落及其紧跟的下一个段落。"""
    return [
        (p, paragraphs[i+1] if i+1 < len(paragraphs) else '')
        for i, p in enumerate(paragraphs) if p.startswith('*')
    ]


def extract_horizontal(folder_path, output_csv):
    docx_files = glob.glob(os.path.join(folder_path, '*.docx'))
    if not docx_files:
        print(f"未找到 docx 文件: {folder_path}")
        return
    all_questions = []
    answers = {}
    for file in docx_files:
        try:
            doc = Document(file)
            paras = [p.text.strip() for p in doc.paragraphs if p.text and p.text.strip()]
            pairs = extract_star_paragraphs(paras)
            name = os.path.basename(file)
            if '】' in name and '【' in name:
                name = name.split('【')[-1].split('】')[0]
            ans = {}
            for q, a in pairs:
                if q not in all_questions:
                    all_questions.append(q)
                ans[q] = a
            answers[name] = ans
        except Exception as e:
            print(f"处理文件出错: {file}, 错误: {e}")
    with open(output_csv, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        writer.writerow(['题目'] + list(answers.keys()))
        for q in all_questions:
            writer.writerow([q] + [answers[n].get(q, '') for n in answers])
    print(f"已横向整合保存 {len(all_questions)} 道题到: {output_csv}")


if __name__ == '__main__':
    folder = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'survey_list'))
    out_csv = os.path.join(os.path.dirname(__file__), 'extracted_pairs.csv')
    extract_horizontal(folder, out_csv)
