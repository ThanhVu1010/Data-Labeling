import pandas as pd
from collections import Counter
import nltk
from nltk.tokenize import word_tokenize
import sys
from nltk.util import ngrams
import re

# Tải tệp Excel
file_path = 'D:\\CongViec\\code\\im_hs_50_2022_converted_date.xlsx'
try:
    df = pd.read_excel(file_path)
except Exception as e:
    print(f"Lỗi khi đọc tệp Excel: {e}")
    df = None

if df is not None:
    # Tải bộ từ điển của NLTK
    nltk.download('punkt')

    # Tạo từ điển sửa lỗi chính tả
    correction_dict = {
        'tămg': 'tằm',
        'tăm': 'tằm',
        'tầm': 'tằm',
    }

    # Hàm để sửa lỗi chính tả sử dụng từ điển tùy chỉnh
    def correct_spelling_with_dict(product):
        if pd.isnull(product):
            return ''

        # Bỏ dấu gạch dưới trong chuỗi sản phẩm
        cleaned_product = str(product).replace('_', ' ')

        # Sử dụng regex để tách các từ và giữ nguyên số thập phân
        pattern = re.compile(r'\b\d+\.\d+\b|\b\w+\b')
        words = pattern.findall(cleaned_product)

        # Sử dụng từ điển để sửa lỗi chính tả
        corrected_words = []
        for word in words:
            # Kiểm tra và sửa các từ trong từ điển
            corrected_word = correction_dict.get(word.lower(), word)
            corrected_words.append(corrected_word)

        # Kết hợp các từ đã sửa lại thành một chuỗi
        corrected_text = ' '.join(corrected_words)

        return corrected_text

    df['Corrected_Products'] = df['Products'].apply(correct_spelling_with_dict)

    # Hàm để tách thành các từ riêng lẻ sử dụng NLTK
    def split_into_words(product):
        words = word_tokenize(str(product).title())  
        words = [word for word in words if word.isalnum()]  
        return len(words), words

    # Áp dụng hàm vào cột 'Corrected_Products'
    df['word_count'], df['words'] = zip(*df['Corrected_Products'].apply(split_into_words))

    # Tạo một DataFrame mới với mỗi từ là một cột riêng lẻ
    words_df = pd.DataFrame(df['words'].tolist(), columns=[f'word_{i}' for i in range(df['words'].apply(len).max())])

    # Kết hợp số lượng từ và các từ đã tách vào một DataFrame
    result_df = pd.concat([df[['Products', 'Corrected_Products', 'word_count']], words_df], axis=1)

    # Làm phẳng danh sách các từ
    all_words = [word for sublist in df['words'].tolist() for word in sublist]

    # Đếm tần suất của mỗi từ
    word_counts = Counter(all_words)

    # Chuyển đổi Counter thành DataFrame
    word_freq_df = pd.DataFrame(word_counts.items(), columns=['Word', 'Frequency'])

    # Sắp xếp DataFrame theo tần suất giảm dần
    word_freq_df = word_freq_df.sort_values(by='Frequency', ascending=False)

    # Hàm để tạo n-gram
    def generate_ngrams(words, n):
        n_grams = ngrams(words, n)
        return [' '.join(grams) for grams in n_grams]

    # Tạo n-gram (ví dụ: bigram, trigram)
    df['bigrams'] = df['words'].apply(lambda x: generate_ngrams(x, 2))
    df['trigrams'] = df['words'].apply(lambda x: generate_ngrams(x, 3))
    df['4-grams'] = df['words'].apply(lambda x: generate_ngrams(x, 4))
    df['5-grams'] = df['words'].apply(lambda x: generate_ngrams(x, 5))
    df['6-grams'] = df['words'].apply(lambda x: generate_ngrams(x, 6))
    # Làm phẳng danh sách các n-gram
    all_bigrams = [bigram for sublist in df['bigrams'].tolist() for bigram in sublist]
    all_trigrams = [trigram for sublist in df['trigrams'].tolist() for trigram in sublist]
    all_fourgrams = [fourgram for sublist in df['4-grams'].tolist() for fourgram in sublist]
    all_fivegrams = [fivegram for sublist in df['5-grams'].tolist() for fivegram in sublist]
    all_sixgrams = [sixgram for sublist in df['6-grams'].tolist() for sixgram in sublist]
    # Đếm tần suất của mỗi n-gram
    bigram_counts = Counter(all_bigrams)
    trigram_counts = Counter(all_trigrams)
    fourgram_counts = Counter(all_fourgrams)
    fivegram_counts = Counter(all_fivegrams)
    sixgram_counts = Counter(all_sixgrams)

    # Chuyển đổi Counter thành DataFrame
    bigram_freq_df = pd.DataFrame(bigram_counts.items(), columns=['Bigram', 'Frequency'])
    trigram_freq_df = pd.DataFrame(trigram_counts.items(), columns=['Trigram', 'Frequency'])
    fourgram_freq_df = pd.DataFrame(fourgram_counts.items(), columns=['4-gram', 'Frequency'])
    fivegram_freq_df = pd.DataFrame(fivegram_counts.items(), columns=['5-gram', 'Frequency'])
    sixgram_freq_df = pd.DataFrame(sixgram_counts.items(), columns=['6-gram', 'Frequency'])

    # Sắp xếp DataFrame theo tần suất giảm dần
    bigram_freq_df = bigram_freq_df.sort_values(by='Frequency', ascending=False)
    trigram_freq_df = trigram_freq_df.sort_values(by='Frequency', ascending=False)
    fourgram_freq_df = fourgram_freq_df.sort_values(by='Frequency', ascending=False)
    fivegram_freq_df = fivegram_freq_df.sort_values(by='Frequency', ascending=False)
    sixgram_freq_df = sixgram_freq_df.sort_values(by='Frequency', ascending=False)

    # Lưu kết quả vào tệp Excel mới
    output_file_path = 'D:\\CongViec\\code\\test_n-gram_im.xlsx'
    with pd.ExcelWriter(output_file_path) as writer:
        result_df.to_excel(writer, sheet_name='Processed Products', index=False)
        word_freq_df.to_excel(writer, sheet_name='Word Frequencies', index=False)
        bigram_freq_df.to_excel(writer, sheet_name='Bigram Frequencies', index=False)
        trigram_freq_df.to_excel(writer, sheet_name='Trigram Frequencies', index=False)
        fourgram_freq_df.to_excel(writer, sheet_name='4-gram Frequencies', index=False)
        fivegram_freq_df.to_excel(writer, sheet_name='5-gram Frequencies', index=False)
        sixgram_freq_df.to_excel(writer, sheet_name='6-gram Frequencies', index=False)
    print(f"Dữ liệu đã được xử lý và lưu vào '{output_file_path}'")
else:
    print("Không thể đọc tệp Excel.")
