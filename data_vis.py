import pandas as pd
import matplotlib
matplotlib.use('TkAgg')
import matplotlib.pyplot as plt
from collections import Counter

# importing the spreadsheet into a dataframe
data = pd.read_excel('Email_Analytics.xlsx')

# Weekly traffic
# sort by day of the week
data['Day'] = pd.Categorical(data['Day'], categories= ['Mon','Tue','Wed','Thu','Fri','Sat', 'Sun'],ordered=True)
count_sorted_by_day = data['Day'].value_counts().sort_index()

plt.figure(1)
count_sorted_by_day.plot(marker = 'o', color = 'blueviolet', linewidth = 2, ylim = [0,750])
plt.title('Weekly Email Traffic', fontweight = 'bold' ,fontsize = 14)
plt.ylabel("Received Email Count", fontweight = 'bold', labelpad = 15)
plt.grid()


# Hourly traffic
# splitting only the hour portion in the time column
# sort by hour of the day - using sort_index for numeric sort
received = data[data['Sent/Received'] == 'Received']
hour = received['Time'].str.split(':').str[0] + ':00'
count_sorted_by_hour = hour.value_counts().sort_index()

plt.figure(2)
count_sorted_by_hour.plot(marker = 'o', color = 'green')
plt.title('Hourly Email Traffic', fontsize = 14, fontweight = 'bold')
plt.ylabel("Received Email Count", fontweight = 'bold', labelpad = 15)
plt.xlabel("Hour of the Day", fontweight = 'bold', labelpad = 15)
plt.xticks(range(len(count_sorted_by_hour.index)), count_sorted_by_hour.index)
plt.xticks(rotation=90)
plt.grid()

# Sent vs. Received
fig = plt.figure(3)
fig.patch.set_facecolor('black')
sent = data[data['Sent/Received'] == 'Sent']
sent_count = sent.shape[0]
received_count = received.shape[0]
values = [sent_count, received_count]
labels = ['Sent', 'Received']
colors = ['red', 'green']

def make_autopct(values):
    def my_autopct(pct):
        total = sum(values)
        val = int(round(pct*total/100.0))
        return '{p:.1f}%  ({v:d})'.format(p=pct,v=val)
    return my_autopct

ax3 = plt.pie(values, labels = labels, colors = colors, textprops={'color':"w", 'fontweight':'bold'}, startangle= 120, autopct=make_autopct(values))
plt.title('Sent vs. Received', fontsize = 14 ,fontweight = 'bold', color = 'white')

# subject word count histogram
# count the number of words in the subject
data['Subject Word Count'] = data['Subject'].str.split(' ').fillna('none').str.len()

plt.figure(4)
plt.hist(data['Subject Word Count'], bins=15, color = 'slategray', ec = 'black')
plt.axis([0, 30, 0, 1200])
plt.xlabel('Word Count', fontweight = 'bold')
plt.ylabel('No. of Emails', fontweight = 'bold')
plt.title('Subject Word Count Histogram', fontsize = 14, fontweight = 'bold')

# top senders horizontal bar chart
sender_top_20 =  received['From (Sender)'].value_counts().nlargest(20)
sender_top_20_count = sender_top_20.values
sender_top_20_names = sender_top_20.index.tolist()

plt.figure(5)
plt.barh(sender_top_20_names, sender_top_20_count, color = 'gold', ec = 'black', linewidth = 1.0)
plt.gca().invert_yaxis()
plt.title('Top 20 Senders', fontsize = 14 ,fontweight = 'bold')
plt.xlabel('Received Email Count', fontweight = 'bold')
plt.tight_layout()


# top words used in subjects
word_list_2d = data['Subject'].str.split(' ').fillna('none').tolist()
word_list_1d = [word for list in word_list_2d for word in list]
word_list_1d = [word.lower() for word in word_list_1d]
exclude_list = ['this', 'that', 'your', 'with', 'from']
word_list_1d = [word for word in word_list_1d if word not in exclude_list and len(word)>3]
common_words_map = Counter(word_list_1d).most_common(10)
common_words = [pair[0] for pair in common_words_map]
frequency = [pair[1] for pair in common_words_map]

plt.figure(6)
plt.barh(common_words, frequency, color = 'lightcoral', ec = 'black', linewidth = 1.25)
plt.gca().invert_yaxis()
plt.title('Most Common Words in Subjects', fontsize = 14 ,fontweight = 'bold')
y = 0.15
for i in range(len(frequency)):
    if len(str(frequency[i])) == 3:
        x = frequency[i] - 14
    else:
        x = frequency[i] - 10
    plt.text(x,y,frequency[i], fontsize = 10,fontweight = 'bold')
    y = y + 1
plt.xticks([0,200])
plt.xlabel('Occurrences', fontweight = 'bold', labelpad=-5)
plt.show()
