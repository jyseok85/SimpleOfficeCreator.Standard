# SimpleOfficeCreator.Stardard
Create Word, PPT, Excel Simply


사용가능한 컨트롤은 4종류입니다.
- 이미지
- 텍스트상자
- 도형(원,네모)
- 테이블

# 파워포인트


# 워드
일반 적인 문서 작성용이 아니기 때문에 텍스트를 기본 배경에 쓰지 않습니다. 
PPT처럼 텍스트 상자를 통해서 텍스트를 입력하는 방법을 사용하게 됩니다. 

<hr/>

# 텍스트 상자
![image](https://github.com/jyseok85/SimpleOfficeCreator.Standard/assets/48501866/b4a8d3b0-3262-41f6-9959-e7ba4940e1ff)
- 필수 : 위치, 크기, 텍스트
- 선택 : 글꼴설정, 단락설정, 도형스타일

   ## 글꼴 설정
   > #### 설정가능한 속성
   > - 폰트명
   > - 폰트사이즈
   > - 스타일(굵게,기울임,밑줄,취소선)
   > - 문자간격
   > - 색상
   >
   > ###### UI 컨트롤 위치
   >|Word or Powerpoint|
   >|:---:|
   >|![image](https://github.com/jyseok85/SimpleOfficeCreator.Standard/assets/48501866/de262be2-24f0-4c97-b8f7-1d85d3028409)|
   > 
   > 
   > |Word|Powerpoint|
   > |:---:|:---:|
   > |![image](https://github.com/jyseok85/SimpleOfficeCreator.Standard/assets/48501866/f0928a0c-e14b-4896-9ad7-176faf384846)|![image](https://github.com/jyseok85/SimpleOfficeCreator.Standard/assets/48501866/0c4dd574-cc15-4fdf-9f71-6b59d169f1ef)|
  > <details><summary>Code</summary>    
  >
  > ``` C#
  > //글꼴을 설정합니다.[옵션]
  > var font = new OfficeFont()
  > {
  >    Name = "맑은 고딕",
  >     Size = 10,
  >     Bold = false,
  >     Italic = false,
  >     UnderLine = false,
  >     Strike = false,
  >     CharacterSpacing = 0,
  >     Color = "black"
  > };
  > ```
  </details>
  
  ## 단락 설정
  > #### 설정가능한 속성
  > - 가로 정렬
  > - 세로 정렬
  > - 텍스트 방향
  > - 줄 간격
  >   
  > ###### UI 컨트롤 위치
  >|Word|Powerpoint|
  >|:---:|:---:|
  >|![image](https://github.com/jyseok85/SimpleOfficeCreator.Standard/assets/48501866/76ca1cca-5212-40b3-a495-6ddb829da3f7)|![image](https://github.com/jyseok85/SimpleOfficeCreator.Standard/assets/48501866/a2c1a90d-5d1f-4092-a2f8-d61e55c2a608)|
  > <details><summary>Code</summary>    
  >
  > ``` C# 
  > //단락을 설정합니다.[옵션]
  > var paragraph = new OfficeParagraph()
  > {
  >    AlignmentHorizontal = TextAlignmentHorizontal.Center,
  >    AlignmentVertical = TextAlignmentVertical.Center,
  >    TextDirection = TextDirection.Stacked,
  >    LineSpacing = 0
  > };
  > ```
  </details>
  
  ## 도형 스타일 설정
  > #### 설정가능한 속성
  > - 도형 채우기 사용유무
  > - 도형 채우기 색
  > - 도형 윤곽선 사용 유무
  > - 도형 윤곽선 색
  > - 도형 윤곽선 두께
  > - 도형 윤곽선 스타일
  > - 도형 모양(네모, 원)
  >   
  > ###### UI 컨트롤 위치
  >|Word or Powerpoint|
  >|:---:|
  >|![image](https://github.com/jyseok85/SimpleOfficeCreator.Standard/assets/48501866/39ecd583-3caf-46a3-bcce-ef05e9452d22)|
  > <details><summary>Code</summary>    
  >
  > ``` C# 
  > //도형 스타일을 설정합니다.[옵션]
  >  var style = new OfficeShapeStyle()
  > {
  >    UseFill = true,
  >    FillColor = "yellow",
  >    UseOutline = true,
  >    OutlineWeight = 1,
  >    OutlineColor = "black",
  >    OutlineDashes = "solid",
  >    ShapeTypeValue = "circle"
  > };
  > ```
  </details>

### 전체코드
``` C#
public string CreateSingleTextBoxDocument()
{
    //생성할 문서를 지정한다.
    var officeCreator = new OfficeCreator(OfficeType.Word);
    //글꼴을 설정합니다.[옵션]
    var font = new OfficeFont()
    {
        Name = "맑은 고딕",
        Size = 10,
        Bold = false,
        Italic = false,
        UnderLine = false,
        Strike = false,
        CharacterSpacing = 0,
        Color = "black"
    };
    //단락을 설정합니다.[옵션]
    var paragraph = new OfficeParagraph()
    {
        AlignmentHorizontal = TextAlignmentHorizontal.Center,
        AlignmentVertical = TextAlignmentVertical.Center,
        TextDirection = TextDirection.Stacked,
        LineSpacing = 0
    };
    var style = new OfficeShapeStyle()
    {
        UseFill = true,
        FillColor = "yellow",
        UseOutline = true,
        OutlineWeight = 1,
        OutlineColor = "black",
        OutlineDashes = "solid",
        ShapeTypeValue = "rectangle"
    };

    //텍스트 박스 생성
    var model = new OfficeModelCreator().CreateTextBox(0, 0, 100, 100, "텍스트 박스 \n테스트",  font, paragraph, style); 
    //모델 목록을 만들고 생성한 컨트롤을 추가한다. 
    var officeModels = new List<OfficeModel>();
    officeModels.Add(model);
    //변환한다. 
    officeCreator.ConvertPage(1, officeModels);
    return officeCreator.Save();
}
```

# 이미지 
- 필수 : 위치, 크기, 이미지 Base64 데이터
- 스타일 지원 : 테두리(색, 두꼐, 모양)  
- ![image](https://github.com/jyseok85/SimpleOfficeCreator.Standard/assets/48501866/01fb019e-3336-4d63-bd7c-456c800e4606)

<details><summary>Code</summary>
  
``` C#
public string CreateOneImage()
{
    var officeCreator = new OfficeCreator(OfficeType.Word);

    //이미지를 가져온다.
    string base64String = Utils.Instance.GetWebImage(
           "https://img-prod-cms-rt-microsoft-com.akamaized.net/cms/api/am/imageFileData/RE1Mu3b?ver=5c31");

    //테두리 스타일 지정(옵션)
    var officePictureStyle = new OfficePictureStyle()
    {
        Weight = 3,
        Color = "blue",
        NoOutline = false,
        Dashes = "DashDotDot"
    };

    //이미지 생성
    var model = new OfficeModelCreator().CreatePicture(0, 0, 100, 100, base64String, officePictureStyle);
    
    //모델 목록을 만들고 이미지를 추가한다. 
    var officeModels = new List<OfficeModel>();
    officeModels.Add(model);

    //변환한다. 
    officeCreator.ConvertPage(1, officeModels);

    return officeCreator.Save();
}
```
</details>




### 테이블

### 도형


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
![image](https://github.com/jyseok85/SimpleOfficeCreator.Stardard/assets/48501866/f747cdd5-b653-4900-b460-bf8a3e1fcd29)

텍스트방향
- 테이블은 특별히 없음.

워드는 스택형이 없네?

PPT랑 워드랑 디자인은 비슷한데 왜 다른거임?
- 텍스트박스에서는 가로, 세로만 지원
