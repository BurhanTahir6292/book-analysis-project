import pandas as pd
import random
from textblob import TextBlob
import matplotlib.pyplot as plt
import seaborn as sns
import os

def generate_mock_review(rating):
    if rating == 5:
        return random.choice([
            "Absolutely amazing! I couldn't put it down.",
            "A masterpiece. Highly recommended to everyone.",
            "Incredible read, totally worth the price."
        ])
    elif rating == 4:
        return random.choice([
            "Great book, really enjoyed it.",
            "Very good read, though the pacing was a bit slow in the middle.",
            "Solid and entertaining. Would read again."
        ])
    elif rating == 3:
        return random.choice([
            "It was okay. Nothing spectacular but not bad.",
            "An average read. Had some good moments."
        ])
    elif rating == 2:
        return random.choice([
            "Disappointing. The plot holes were too obvious.",
            "Not my cup of tea. Felt dragged out."
        ])
    else: # rating 1
        return random.choice([
            "Terrible book. Do not waste your time.",
            "Awful. The worst thing I have ever read.",
            "Extremely boring and poorly written. A complete waste of money."
        ])

def analyze_sentiment(text):
    blob = TextBlob(text)
    polarity = blob.sentiment.polarity
    if polarity > 0.1:
        return 'Positive'
    elif polarity < -0.1:
        return 'Negative'
    else:
        return 'Neutral'

def detect_emotion(text, rating):
    text_lower = str(text).lower()
    
    joy_words = ['masterpiece', 'recommend', 'great', 'enjoyed', 'liked', 'fantastic', 'best', 'solid', 'entertaining', 'good', 'wonderful', 'loved', 'love', 'superb']
    surprise_words = ['amazing', 'incredible', 'wonder', 'surprised']
    disappointment_words = ['disappointing', 'disappointment', 'expected a bit more', 'below average', 'letdown']
    sadness_words = ['dull', 'slow', 'dragged out', 'struggled', 'boring', 'sad']
    anger_words = ['terrible', 'awful', 'worst', 'waste', 'hated', 'nonsensical', 'zero redeeming', 'poorly written']

    score_joy = sum(text_lower.count(w) for w in joy_words)
    score_surprise = sum(text_lower.count(w) for w in surprise_words)
    score_disappointment = sum(text_lower.count(w) for w in disappointment_words)
    score_sadness = sum(text_lower.count(w) for w in sadness_words)
    score_anger = sum(text_lower.count(w) for w in anger_words)

    if rating == 5:
        score_joy += 2
        score_surprise += 1
    elif rating == 4:
        score_joy += 1
    elif rating == 2:
        score_disappointment += 1
        score_sadness += 1
    elif rating == 1:
        score_anger += 2
        score_disappointment += 1

    scores = {
        'Joy': score_joy,
        'Surprise': score_surprise,
        'Disappointment': score_disappointment,
        'Sadness': score_sadness,
        'Anger': score_anger
    }

    max_emotion = max(scores, key=scores.get)
    if scores[max_emotion] == 0:
        return 'Neutral'
    return max_emotion

def perform_sentiment_and_emotion_analysis(df):
    """
    Task 4: Sentiment and Emotion Analysis
    Applies NLP to text features to derive sentiment and specific emotions.
    """
    print("Running NLP Sentiment and Emotion Analysis...")
    
    # Generate mock reviews if none exist
    if 'Review_Text' not in df.columns:
        df['Review_Text'] = df['Numeric Rating'].apply(generate_mock_review)
        
    # Apply sentiment analysis using TextBlob
    df['Sentiment_Polarity'] = df['Review_Text'].apply(lambda x: TextBlob(x).sentiment.polarity)
    df['Sentiment_Class'] = df['Review_Text'].apply(analyze_sentiment)
    
    # Apply emotion detection based on lexicon and rating
    df['Emotion'] = df.apply(lambda row: detect_emotion(row['Review_Text'], row['Numeric Rating']), axis=1)
    
    # Chart: Sentiment Distribution
    sns.set_theme(style="whitegrid")
    plt.figure(figsize=(8, 6))
    sentiment_counts = df['Sentiment_Class'].value_counts()
    sns.barplot(x=sentiment_counts.index, y=sentiment_counts.values, hue=sentiment_counts.index, palette={'Positive':'green', 'Neutral':'gray', 'Negative':'red'}, legend=False)
    plt.title('Sentiment Distribution of Book Reviews', fontsize=14, fontweight='bold')
    plt.xlabel('Sentiment', fontsize=12)
    plt.ylabel('Count', fontsize=12)
    plt.tight_layout()
    
    assets_dir = "charts_output"
    os.makedirs(assets_dir, exist_ok=True)
    chart_path = os.path.join(assets_dir, 'chart_sentiment.png')
    plt.savefig(chart_path, dpi=300)
    plt.close()
    print(f"Saved: {chart_path}")
    
    return df

if __name__ == "__main__":
    # Create mock data to test the NLP script standalone
    mock_data = {
        'Title': ['Book A', 'Book B', 'Book C'],
        'Numeric Rating': [5, 3, 1]
    }
    df_mock = pd.DataFrame(mock_data)
    analyzed_df = perform_sentiment_and_emotion_analysis(df_mock)
    print("\n--- Output with Sentiment and Emotion ---")
    print(analyzed_df[['Title', 'Review_Text', 'Sentiment_Class', 'Emotion']])
