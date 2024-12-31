import pandas as pd
import nltk
from nltk.corpus import stopwords
from nltk.stem import WordNetLemmatizer
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.model_selection import train_test_split
from sklearn.multioutput import MultiOutputClassifier
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import classification_report
from sklearn.preprocessing import MultiLabelBinarizer
import numpy as np

# Download necessary NLTK data files (uncomment if needed)
# nltk.download('stopwords')
# nltk.download('punkt')
# nltk.download('wordnet')

# Initialize stop words and lemmatizer
stop_words = set(stopwords.words('english'))
lemmatizer = WordNetLemmatizer()

# Expanded sample data
data = pd.DataFrame({
    'statement': [
        "AI and machine learning are transforming industries",
        "Data science uses statistical methods to analyze data",
        "Azure provides cloud services for businesses",
        "Deep learning models are a subset of AI and ML",
        "Data visualization is key to communicating data insights",
        "Cloud computing platforms like AWS and Azure are in high demand",
        "AI and data science are reshaping the future of technology",
        "Big data analytics helps organizations to make data-driven decisions",
        "Machine learning algorithms are used in predictive analytics",
        "The future of cloud computing is in hybrid and multi-cloud solutions",
        "Natural Language Processing is a subfield of AI focused on language data",
        "Business intelligence (BI) tools help companies analyze business data",
        "Cybersecurity is important in the age of digital transformation",
        "DevOps practices are essential for continuous integration in software development"
    ],
    'tags': [
        ["AI", "ML", "Technology"],
        ["Data Science", "Analysis", "Statistics"],
        ["Cloud", "Azure", "Services"],
        ["AI", "ML", "Deep Learning"],
        ["Data Science", "Visualization"],
        ["Cloud", "AWS", "Azure", "Demand"],
        ["AI", "Data Science", "Technology"],
        ["Big Data", "Analytics", "Data Science"],
        ["Machine Learning", "Algorithms", "Predictive Analytics"],
        ["Cloud", "Computing", "Hybrid", "Multi-cloud"],
        ["AI", "NLP", "Technology"],
        ["Business Intelligence", "Data", "Analysis"],
        ["Cybersecurity", "Digital Transformation"],
        ["DevOps", "Software Development", "CI/CD"]
    ]
})

# Preprocess statements with NLTK
def preprocess_text(text):
    tokens = nltk.word_tokenize(text.lower())
    tokens = [lemmatizer.lemmatize(word) for word in tokens if word.isalnum() and word not in stop_words]
    return ' '.join(tokens)

# Apply preprocessing to the 'statement' column
data['processed_statement'] = data['statement'].apply(preprocess_text)

# Display the data after preprocessing
print("Processed Data:")
print(data)

# Vectorize the preprocessed statements using TF-IDF
vectorizer = TfidfVectorizer()
X = vectorizer.fit_transform(data['processed_statement'])

# Encode the tags (multi-label binarization)
mlb = MultiLabelBinarizer()
y = mlb.fit_transform(data['tags'])

# Check label distribution in the full data
label_counts = np.sum(y, axis=0)
print("\nLabel distribution in the full dataset:")
print(label_counts)  # Sum of labels per class across all samples

# Identify labels with only 1 instance
rare_labels = np.where(label_counts == 1)[0]
print("\nRare labels (appearing only once):", rare_labels)

# Manually ensure that rare labels are in the training set
train_index = []
test_index = []

# Manually assign rows with rare labels to the training set
for i in range(len(y)):
    if np.any(y[i, rare_labels]):  # If the row contains any rare label
        train_index.append(i)
    else:
        test_index.append(i)

# Create the train-test split based on the manual assignment
X_train, X_test = X[train_index], X[test_index]
y_train, y_test = y[train_index], y[test_index]

# Check label distribution in the training data
print("\nLabel distribution in the training set:")
print(np.sum(y_train, axis=0))  # Sum of labels per class in the training set

# Check if any class is missing in the training set
if np.any(np.sum(y_train, axis=0) == 0):
    print("\nWarning: Some classes are missing from the training set!")

# Train the model using Logistic Regression with MultiOutputClassifier
model = MultiOutputClassifier(LogisticRegression(max_iter=1000, solver='lbfgs'))
model.fit(X_train, y_train)

# Predict and evaluate the model
y_pred = model.predict(X_test)

# Display the classification report
report = classification_report(y_test, y_pred, target_names=mlb.classes_)
print("\nClassification Report:")
print(report)
