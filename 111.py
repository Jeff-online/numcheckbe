import io
import fitz
import regex as re


def re_replace(content):
    # total_re = re.findall(r"(?<!20)\d{2}年度|※?(.*?※)|加速|\d{1,2}～\d{1,2}月期|\d{1,2}月\d{1,2}日～\d{1,2}月\d{1,2}日|マイナスに寄与", content)
    year_re = re.findall(r"\d{2,4}年度", content)
    symbol_re = re.findall(r"※?(.*?※)", content)
    speed_re = re.findall(r"加速", content)
    day_re = re.findall(r"\d{1,2}[~～]\d{1,2}月期|\d{1,2}月\d{1,2}日[~～]\d{1,2}月\d{1,2}日", content)
    word_re = re.findall(r"マイナスに寄与|サステナブル|エンターテイメント|外国人投資家からの資金流入|外国人投資家の資金流出|魅力|投資妙味|割高|割高感|割安|割安感|MSCIインド指数|ダウ平均|NYダウ|への組み入れ", content)      # サステナブル　サスティナブル   |   エンターテイメント　エンターテインメント
    score_re = re.findall(r"[\d.]+?[～~][\d.]+?[%％]", content)
    half_re = re.findall(r"\d{2,4}年第[1-4一二三四]四半期", content)
    print(year_re, symbol_re, speed_re, day_re, word_re, score_re, half_re)


def add_comments_to_pdf(pdf_bytes, corrections):
    """
    PDF 파일에서 틀린 부분을 찾아 코멘트를 추가합니다.

    :param pdf_bytes: PDF 파일의 바이트 데이터
    :param corrections: 수정 사항 리스트 (각 항목은 page, original_text, comment를 포함)
    :return: 수정된 PDF 파일의 BytesIO 객체
    """
    # 입력 유효성 검사
    if not isinstance(pdf_bytes, bytes):
        raise ValueError("pdf_bytes must be a bytes object.")
    if not isinstance(corrections, list):
        raise ValueError("corrections must be a list of dictionaries.")
    for correction in corrections:
        if not all(key in correction for key in ["page", "original_text", "comment"]):
            raise ValueError("Each correction must contain 'page', 'original_text', and 'comment' keys.")

    try:
        # PDF 파일 열기 (BytesIO에서 직접 열기)
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    except Exception as e:
        raise ValueError(f"Invalid PDF file: {str(e)}")

    for idx, correction in enumerate(corrections):
        page_num = correction["page"]
        comment = correction["comment"]
        reason_type = correction["reason_type"]
        locations = correction["locations"][0]
        text_instances = [fitz.Rect(locations["x0"], locations["y0"], locations["x1"], locations["y1"])]
        if int(text_instances[0][0]) == 0:
            continue

        # 페이지 번호 유효성 검사
        if page_num < 0 or page_num >= len(doc):
            raise ValueError(f"Invalid page number: {page_num}")

        page = doc.load_page(page_num)

        for rect in text_instances:
            highlight = page.add_rect_annot(rect)
            highlight.set_colors(stroke=None, fill=(1, 1, 0))
            highlight.set_opacity(0.5)
            highlight.set_info({
                "title": reason_type,  # 可选：显示在注释框标题栏
                "content": comment
            })
            highlight.update()

    # 수정된 PDF를 BytesIO에 저장
    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    doc.close()

    return output


if __name__ == '__main__':
    # re_replace("※23年度GDP<sup>※<sup>は2023年全年国民年収指数※の英語表現です。去年の景気上昇により、明らかな加速が示されていました。特に23年第4四半期※と24年4~6月期の上昇は30%以上となり、市場経済にマイナスに寄与しました。3年生と五年生は3～20％の成長が見えました。")

    # print(list(a()))
    # input_text = "-1～0.5％"
    # score_re = re.findall(r"[\d.]+?[～~][\d.]+?[%％]", input_text)
    # for score_result in score_re:
    #     cor_score = score_result.replace("～", "％～").replace("~", "％~")
    #     print(cor_score)
    import pandas as pd
    from io import StringIO

    content = "〇フランス国債を高位に組み入れ、フランス国債ポートフォリオの平均残存期間を9～10年程度に維持しました。"

    category = "比率"
    condition = "{\"1\":{\"2\":\"ポートフォリオ特性値\",\"4\":\"平均クーポン\",\"5\":\"平均直利\",\"6\":\"平均最終利回り\",\"7\":\"平均残存期間\",\"8\":\"平均デュレーション\",\"9\":\"フランス国債\",\"10\":\"ドイツ短期国債先物\"},\"2\":{\"2\":null,\"4\":0.0266608436,\"5\":0.0251916388,\"6\":0.0286821519,\"7\":9.6321311392,\"8\":-0.0109588284,\"9\":7.5287110205,\"10\":-7.5396698489},\"3\":{\"2\":null,\"4\":null,\"5\":null,\"6\":null,\"7\":\"←小数点第１位まで表示\",\"8\":null,\"9\":null,\"10\":null}}"

    if condition:
        result_temp = []
        table_list = condition.split("\n")
        for data in table_list:
            if data:
                if category in ["比率", "配分"]:
                    re_num = re.search(r"([-\d. ]+)(%|％)", content)
                    if re_num:
                        num = re_num.groups()[0]
                        float_num = len(str(num).split(".")[1]) if "." in num else 0
                        old_data = pd.read_json(StringIO(data))
                        result_temp.append(old_data.applymap(
                            lambda x: (str(round(x * 100, float_num)) + "%" if float_num != 0 else str(
                                int(round(x * 100, float_num))) + "%")
                            if not pd.isna(x) and isinstance(x, float) else x).to_json(force_ascii=False))
                    else:
                        result_temp.append(pd.read_json(StringIO(data)).to_json(force_ascii=False))
                else:
                    result_temp.append(pd.read_json(StringIO(data)).to_json(force_ascii=False))
        if len(result_temp) > 1:
            for i in result_temp:
                # print(pd.read_json(i))
                print(i)
                break
            result_data = "\n".join(result_temp)
        else:
            result_data = result_temp[0]
            print(pd.read_json(result_data))
    else:
        result_data = ""

    # print(result_data)
