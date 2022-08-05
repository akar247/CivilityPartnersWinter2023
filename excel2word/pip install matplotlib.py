pip install matplotlib
Options = ["Strongly disagree", "Disagree", "Agree", "Strongly agree"]
Responses = [14.29%,  11.90%, 50.00%, 23.81%]

plt.bar(xAxis,yAxis)
plt.title('title name')
plt.xlabel('xAxis name')
plt.ylabel('yAxis name')
plt.show()
from http.client import responses
import matplotlib.pyplot as plt
   
Options = ["Strongly disagree", "Disagree", "Agree", "Strongly agree"]
Responses = [14.29%,  11.90%, 50.00%, 23.81%]

plt.bar(Options, Responses)
plt.title('Options vs Responses')
plt.xlabel('Options')
plt.ylabel('Responses')
plt.show()
import matplotlib.pyplot as plt
   
Options = ["Strongly disagree", "Disagree", "Agree", "Strongly agree"]
Responses = [14.29%,  11.90%, 50.00%, 23.81%]

New_Colors = ['green','blue','purple','brown',]
plt.bar(Options, Responses, color=New_Colors)
plt.title('Options vs Responses', fontsize=14)
plt.xlabel('Options', fontsize=14)
plt.ylabel('Responses', fontsize=14)
plt.grid(True)
plt.show()