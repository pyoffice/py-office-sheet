# workbench
A cross-platform spreadsheet based on pandas and numpy for effecient data processing

in progress...

- default saving format: pandas data object(.pdobj), upto 3 times less memory use compare to generic format. this is done by joblib, therefore any application that has python support would be able to read it. It makes data managment easier when doing data analyze or machine learning.
```
   import joblib
   pandas_data_frame = joblin.load("yourfile.pdobj")
```
- support interactive python command
- support python scripting

light mode:
![alt text](https://raw.githubusercontent.com/YC-Lammy/np_spreadsheet/main/doc/Screenshot_20210609_111555.png)


dark mode:
![alt text](https://raw.githubusercontent.com/YC-Lammy/np_spreadsheet/main/doc/Screenshot_20210608_145022.png)