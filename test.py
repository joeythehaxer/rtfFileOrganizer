from langdetect import detect
import re


def remove_non_english(text):
    # Tokenize the text into words
    words = re.findall(r'\b\w+\b', text)

    # Filter out non-English words
    english_words = [word for word in words if detect(word) == 'en']

    # Join the English words into a sentence
    english_sentence = ' '.join(english_words)

    return english_sentence


def summarize(text):
    # Remove non-English text
    english_text = remove_non_english(text)

    # Return the entire English text
    return english_text


def rephrase(text):
    # Split the text into sentences
    sentences = re.split(r'(?<!\w\.\w.)(?<![A-Z][a-z]\.)(?<=\.|\?)\s', text)

    # Rephrase each sentence
    rephrased_sentences = []
    for sentence in sentences:
        # Split the sentence into words
        words = sentence.split()

        # Remove non-English words
        english_words = [word for word in words if detect(word) == 'en']

        # Join the English words into a sentence
        rephrased_sentence = ' '.join(english_words)

        rephrased_sentences.append(rephrased_sentence)

    # Join the rephrased sentences into a single string
    rephrased_text = ' '.join(rephrased_sentences)

    return rephrased_text


# Input text
input_text = "FDI 19315/19317/19318/19319/19320/82620/82622/82621/82623/82624/82625 - FIT OF DOOR/FRAME/ DAMAGE TO DOOR/HINGES/ STRIPS AND SEALS  - FLOOR 5 OPPOSITE FLAT 20"

# Remove non-English text and summarize
english_text = summarize(input_text)

# Rephrase the text
rephrased_text = rephrase(english_text)

# Print the rephrased text
print(rephrased_text)
