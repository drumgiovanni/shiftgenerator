# shiftgenerator
バイト先で使える、シフトを自動で組んでくれるプログラムです。

## 使い方

起動すると翌月のシフトを組むことができます。
以下の手順で従業員の情報を入力していってください。

1.　従業員の名前を入力  

2.　従業員の属性を入力  
    　フルタイム（平日勤務）か土日を入力。  

3.　従業員の休み希望を入力  
    　土日勤務従業員の分は土日の中で休みたい日程を入力。   
    　休み希望がない場合は空白のままEnterを入力。

4.　他に従業員がいるのか入力  
    　他に従業員がいない場合は"no"  
    　いる場合は"yes"、あるいはEnterキーを入力。  

![result](https://github.com/drumgiovanni/shiftgenerator/blob/master/gif2.mov.gif)



全従業員分の情報を記入したのち、上の4でnoを入力すると、プログラムが入っているディレクトリにExcelシートが生成される。
そこに完成したシフト、および各従業員の勤務可能日、出勤日が入っている。


![result](https://github.com/drumgiovanni/shiftgenerator/blob/master/gif3.mov.gif)
