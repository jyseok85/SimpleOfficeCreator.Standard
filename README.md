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

