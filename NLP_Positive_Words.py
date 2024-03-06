import pandas as pd
import nltk
import re

def analyze_text(text, positive_words):
    # Regular Expression
    tokens = re.findall(r"\w+", text.lower())

    # Vocabulary
    pattern = r'\w[\w\',-]*\w'
    tokens = nltk.regexp_tokenize(text.lower(), pattern)
    word_dist = nltk.FreqDist(tokens)

    # Stop Words and Word Filtering
    stop_words = set(["i", "me", "my", "myself", "we", "our", "ours", "ourselves", "you", "your", "yours", "yourself", "yourselves", "he", "him", "his", "himself", "she", "her", "hers", "herself", "it", "its", "itself", "they", "them", "their", "theirs", "themselves", "what", "which", "who", "whom", "this", "that", "these", "those", "am", "is", "are", "was", "were", "be", "been", "being", "have", "has", "had", "having", "do", "does", "did", "doing", "a", "an", "the", "and", "but", "if", "or", "because", "as", "until", "while", "of", "at", "by", "for", "with", "about", "against", "between", "into", "through", "during", "before", "after", "above", "below", "to", "from", "up", "down", "in", "out", "on", "off", "over", "under", "again", "further", "then", "once", "here", "there", "when", "where", "why", "how", "all", "any", "both", "each", "few", "more", "most", "other", "some", "such", "no", "nor", "not", "only", "own", "same", "so", "than", "too", "very", "s", "t", "can", "will", "just", "don", "should", "now", "d", "ll", "m", "o", "re", "ve", "y", "ain", "aren", "couldn", "didn", "doesn", "hadn", "hasn", "haven", "isn", "ma", "mightn", "mustn", "needn", "shan", "shouldn", "wasn", "weren", "won", "wouldn", "covid-19", "virus"])
    filtered_dict = {word: word_dist[word] for word in word_dist if word not in stop_words}
    sorted_tokens = sorted(filtered_dict.items(), key=lambda item: -item[1])

    # Positive Tokens
    positive_tokens = [token for token in tokens if token in positive_words]

    return positive_tokens

def main():
    # Load positive words
    with open("positive-words.txt", 'r') as f:
        positive_words = [line.strip() for line in f]
    
    # User input for the Excel file name, sheet names, and columns to analyze
    file_name = input("Enter the Excel file name: ")
    sheets = input("Enter the sheet names separated by a comma: ").split(",")
    columns = input("Enter the column names for each sheet separated by a comma: ").split(",")
    
    # Initialize a DataFrame to collect all results
    all_results = pd.DataFrame()

    for sheet, column in zip(sheets, columns):
        df = pd.read_excel(file_name, sheet_name=sheet.strip())
        if column.strip() in df.columns:
            df[column.strip()+'_NLP'] = df[column.strip()].apply(lambda x: analyze_text(x, positive_words))
            all_results = pd.concat([all_results, df[[column.strip(), column.strip()+'_NLP']]], ignore_index=True)
        else:
            print(f"Column {column.strip()} not found in sheet {sheet.strip()}.")

    # Save the compiled results to a CSV file
    all_results.to_csv(f"{file_name.split('.')[0]}_NLP.csv", index=False)
    print("Results saved to CSV.")

if __name__ == "__main__":
    main()
    
