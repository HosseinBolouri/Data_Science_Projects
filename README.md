# Data Science, Machine Learning & Data Analysis Projects

Data Science portfolio showcasing SQL, Python (Pandas, SciPy, stats, numpy, matplotlib, TensorFlow, scikit-learn), Machine Learning (ML) and Data Analysis projects from DataCamp, Coursera, and independent data science work.

## Projects
### 1. Univariate Linear Regression: City Profit Prediction (Supervised Learning)
- Implements linear regression from scratch (manual cost and gradient) to predict company profit based on city population (dataset: ex1data1.txt).
- Uses gradient descent (alpha=0.01, iterations=1500), plots training data and fitted line, and predicts profit for populations of 35,000 and 70,000.
- Notebook: [Univariate Linear Regression City Profit Prediction.ipynb](Univariate%20Linear%20Regression%20City%20Profit%20Prediction.ipynb)

### 2. Logistic Regression: Exam Admission Classification (Supervised Learning)
- Classifies student admission (0/1) from two exam scores using logistic regression from scratch (sigmoid, unregularized cost & gradient, gradient descent, decision boundary, accuracy).
- Notebook: [Logistcs_Regression_Students.ipynb](Logistcs_Regression_Students.ipynb)

### 3. Regularized Logistic Regression: Microchip Quality Classification
- Predicts accepted vs rejected microchips using two test scores with polynomial feature mapping (degree 6) and L2 regularization.
- Implements feature mapping, regularized cost/gradient, gradient descent, non-linear decision boundary, and training accuracy (dataset: ex1data3.txt).
- Notebook: [Logistic_Regression_Microchips.ipynb](Logistic_Regression_Microchips.ipynb)

### 4. Multi-Class Neural Network: Handwritten Digit Classification
- Dense neural network (25 -> 15 -> 10 logits) for 0–9 digit recognition on 20x20 grayscale images (flattened to 400 features).
- Uses TensorFlow: SparseCategoricalCrossentropy (from_logits), Adam optimizer, loss curve, sample predictions, accuracy via error count.
- Notebook: [Multi_Class_Neural_Network.ipynb](Multi_Class_Neural_Network.ipynb)

### 5. Tree Ensembles: Heart Disease Prediction
- Dataset: Kaggle Heart Failure Prediction (11 original clinical + categorical features → expanded via one-hot encoding).
- Preprocessing: Pandas one-hot encoding (Sex, ChestPainType, RestingECG, ExerciseAngina, ST_Slope).
- Models: Decision Tree, Random Forest, XGBoost (gradient boosting).
- Key tuning:
  - Depth / min_samples_split to control variance (tree).
  - n_estimators, depth for Random Forest (bias–variance balance).
  - Early stopping (eval_set + early_stopping_rounds=10) for XGBoost to prevent overfitting; best_iteration tracked.
- Evaluation: Accuracy comparison across models; ensembles improved generalization vs single tree.
- Insight: Structural limits (max_depth), averaging (Random Forest) and staged boosting + early stopping (XGBoost) progressively reduce overfitting.
- Notebook: [Trees_Ensemble.ipynb](https://github.com/HosseinBolouri/data-science-ml-projects/blob/main/Trees_Ensemble.ipynb)

### 6. Investigating Netflix Data for 90s
- Exploratory data analysis of Netflix movies from the 1990s.
- Notebook: [Investigating_Netflix_Original.ipynb](Investigating_Netflix_Original.ipynb)

### 7. International Students Mental Health Analysis (SQL Project)
- PostgreSQL analysis on impact of stay duration on mental health.
- Notebook: [Project_SQL_Original_Students_Mental_Health.ipynb](Project_SQL_Original_Students_Mental_Health.ipynb)

### 8. NYC Public School SAT Performance Analysis
- Pandas analysis to identify top math schools, overall best schools, borough variability.
- Notebook: [NYC_Project_DataCamp.ipynb](NYC_Project_DataCamp.ipynb)

### 9. Two-Way ANOVA: Braking Performance Analysis
- Two-way ANOVA using SciPy and Pandas.
- Notebook: [ANOVA 2 ways- braking example.ipynb](ANOVA%202%20ways-%20braking%20example.ipynb)

## About Me
- Data Scientist with experience in analytics, statistical modeling, machine learning, and data analysis projects.
- M.Sc. Industrial Engineering - Data Analytics Specialization (Politecnico di Milano & KIT, Jul 2024).
- Focused on applying Data Science, Data Analysis and ML to real-world problems.

## Tech Stack Highlights
- Python (Pandas, NumPy, SciPy, Matplotlib, TensorFlow, scikit-learn)
- SQL (PostgreSQL)
- Statistical analysis & ML algorithms

## Connect
- LinkedIn: https://www.linkedin.com/in/hossein-bolouri
- Kaggle: https://www.kaggle.com/hosseinbolouri1996
