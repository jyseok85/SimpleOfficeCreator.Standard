# SimpleOfficeCreator.Stardard
Create Word, PPT, Excel Simply

목표
사용방법
``` c#
OfficeCreator oc = new OfficeCreator();
OfficeOption oo = new OfficeOption();
//모델들을 만든다.
oc.Model.CreateLabel(x, y, width, height, text, options); 
oc.Model.CreateTable(x, y, width, height, tableInfo);
//변환시킨다.
oc.Convert(type);
//저장한다.
oc.Save(path) ;
```

1. OfficeModel 을 만든다.
  - OfficeCreator.Model.Creator... 를 통해서 만들었다면 내부에 리스트로 생성되어 있다.
2. List<OfficeModel> 을 변환시킨다.


![image](https://github.com/jyseok85/SimpleOfficeCreator.Stardard/assets/48501866/d2e28df7-0975-4f58-98de-95f86d03f39b)
![image](https://github.com/jyseok85/SimpleOfficeCreator.Stardard/assets/48501866/a02d33dd-8712-4134-8a54-edffffef783f)

Word 테이블 셀 텍스트 자동 맞춤
![image](https://github.com/jyseok85/SimpleOfficeCreator.Stardard/assets/48501866/d7a3e48a-9406-4ba1-beb4-d699e82eee79)
