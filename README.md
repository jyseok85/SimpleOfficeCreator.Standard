# SimpleOfficeCreator.Standard
Create Word, PPT, Excel(Not yet) Simply

---

> **Note**  
> 목적 : 
> 기존 워드, 파워포인트 변환 오픈소스 사용법이 너무 어려워서 아~~주 간단한 핵심기능만 넣어서 사용가능하도록 만들었습니다.  
> 당연히 모든 속성이 완벽하지도 않고, 문서속성 같은것도 다 빼버렸습니다. 
> 
> 그러나, 일반 사용자 입장에서는 딱히 문제는 없을거라 생각됩니다. 
---

> **Warning**  
> 개발하면서 테스트 했을때, 보통 컬러값을 잘못 넣을 경우 문서를 열수 없는 오류가 발생됩니다. 

---
> **Info**  
> 1. 일반 적인 문서 작성용이 아니기 때문에 워드의 경우 텍스트를 기본 배경에 쓰지 않습니다.   
>    파워포인트 처럼 텍스트 상자를 통해 입력하는 방법을 사용합니다.
> 2. 명칭이 일관적이지 않습니다. 예를들어 사용유무값이 어떤건 UseFill, Draw, NoOutline 이렇게 나눠집니다.
>    원인은 OpenXml의 속성명을 그대로 사용했기 때문입니다.
>    (아마도 OpenXml 또한 Office 디자인에서 사용되는 명칭이 달라서 저렇게 만든게 아닐까 합니다.)
---

사용가능한 컨트롤은 4종류입니다.
- 이미지
- 텍스트상자
- 도형(원,네모)
- 테이블



<hr/>

# 텍스트 상자
- 필수 : 위치, 크기, 텍스트
- 선택 : 글꼴설정, 단락설정, 도형스타일
   ||
   |:---:|
   |![image](https://github.com/jyseok85/SimpleOfficeCreator.Standard/assets/48501866/b4a8d3b0-3262-41f6-9959-e7ba4940e1ff)|

   ## 글꼴 설정
   > #### 설정가능한 속성
   > - 폰트명
   > - 폰트사이즈
   > - 스타일(굵게,기울임,밑줄,취소선)
   > - 문자간격
   > - 색상
   >
   > ###### UI 컨트롤 위치
   >||
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
  >||
  >|:---:|
  >|![image](https://github.com/jyseok85/SimpleOfficeCreator.Standard/assets/48501866/39ecd583-3caf-46a3-bcce-ef05e9452d22)|
  > <details><summary>Code</summary>    
  >
  > ``` C# 
  > //도형 스타일을 설정합니다.[옵션]
  > var style = new OfficeShapeStyle()
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
<details><summary><h3>전체코드</h3></summary>    

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
    //도형 스타일을 설정합니다.[옵션]
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
### </details>

<hr/>

# 이미지 
- 필수 : 위치, 크기, 이미지 Base64 데이터
- 선택 : 그림 스타일
  
  |Word|Powerpoint|
  |:---:|:---:|
  |![image](https://github.com/jyseok85/SimpleOfficeCreator.Standard/assets/48501866/9f626aae-3b60-404c-9df1-c39d1109e893)|![image](https://github.com/jyseok85/SimpleOfficeCreator.Standard/assets/48501866/a698a326-83af-4a27-b92a-5bfe2bbbd450)|




  ## 그림 스타일 설정
  > #### 설정가능한 속성
  > - 그림 테두리 사용 유무
  > - 그림 테두리 색
  > - 그림 테두리 두께
  > - 그림 테두리 스타일  
  >   
  > ###### UI 컨트롤 위치
  >|![image](https://github.com/jyseok85/SimpleOfficeCreator.Standard/assets/48501866/39ecd583-3caf-46a3-bcce-ef05e9452d22)|
  >|:---:|
  >|![image](https://github.com/jyseok85/SimpleOfficeCreator.Standard/assets/48501866/01fb019e-3336-4d63-bd7c-456c800e4606)|
  > <details><summary>Code</summary>    
  >
  > ``` C#  
  > var officePictureStyle = new OfficePictureStyle()
  > {
  >    Weight = 3,
  >    Color = "blue",
  >    NoOutline = false,
  >    Dashes = "DashDotDot"
  >    UseFill = true,
  > };
  > ```
  </details>
  
<details><summary><h3>전체코드</h3></summary>    
  
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

<hr/>

# 도형
- 필수 : 위치, 크기
- 선택 : 도형 스타일(텍스트 박스의 도형 스타일과 동일합니다.)
  
  |Word|Powerpoint|
  |:---:|:---:|
  |![image](https://github.com/jyseok85/SimpleOfficeCreator.Standard/assets/48501866/ef1e91ee-286b-4c7b-9ca7-f4b6239cced8)|![image](https://github.com/jyseok85/SimpleOfficeCreator.Standard/assets/48501866/b094c0fb-62fc-471b-903b-82757477136c)|

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
  >||
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
  
<details><summary><h3>전체코드</h3></summary>    
  
``` C#
public string CreatesingleShapeDocument()
{
    var officeCreator = new OfficeCreator(OfficeType.Word);

    //도형 스타일을 설정합니다.[옵션]
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

    //도형 생성
    var model = new OfficeModelCreator().CreateShape(50, 30, 100, 100, style);
    
    //모델 목록을 만들고 이미지를 추가한다. 
    var officeModels = new List<OfficeModel>();
    officeModels.Add(model);

    //변환한다. 
    officeCreator.ConvertPage(1, officeModels);

    return officeCreator.Save();
}
```
</details>

<hr/>

# 테이블
- 필수 : 위치, 크기, 행수, 열수 (or 행 높이 목록, 열 너비 목록)
- 선택 : 글꼴설정, 단락설정, 테이블스타일, 이미지 아이디 - (글꼴설정과 단락설정은 텍스트박스와 동일합니다)
  
  |Word|Powerpoint|
  |:---:|:---:|
  |![image](https://github.com/jyseok85/SimpleOfficeCreator.Standard/assets/48501866/424f5268-b964-40d7-aaba-1c9563a6e09e)|![image](https://github.com/jyseok85/SimpleOfficeCreator.Standard/assets/48501866/25e4020e-09e0-439f-806d-4b9246878947)|



  ## 테이블 스타일 설정
  > #### 설정가능한 속성
  >- 음영 사용유무(왜 오피스에서 명칭이 다 다른지 모르겠네요..)
  >- 음영 컬러
  >- 4방향 테두리
  >> - 테두리 사용유무
  >> - 테두리 색
  >> - 테두리 두께
  >> - 테두리 스타일
  >   
  > ###### UI 컨트롤 위치
  >||
  >|:---:|
  >|![image](https://github.com/jyseok85/SimpleOfficeCreator.Standard/assets/48501866/377bc3c7-5d4c-45ce-a9bc-ae769c858e4a)|
  > <details><summary>Code</summary>    
  >
  > ``` C# 
  > //테이블 스타일을 설정합니다.[옵션]
  > var style = new OfficeTableStyles()
  > {
  >     UseShading = true,
  >     ShadingColor = "yellow"
  > };
  > var border = new Modules.Model.Component.TableDesignTab.Border()
  > {
  >     Draw = true, 
  >     Color = "red",
  >     Weight = 1,
  >     Dashes = "solid"
  > };
  > //상하단을 적용합니다.
  > style.Bottom = border;
  > style.Top = border;
  > ```
  </details>

  ## 테이블 생성 예제
  > 1. 기본
  > ``` C#
  > //일반적인 테이블을 생성합니다. 
  > var table1 = new OfficeModelCreator().CreateTable(50, 30, 300, 300, 2, 2);
  > officeModels.Add(table1);
  > new OfficeModelCreator().CreateTableCell(table1, 0, 0, "PowerPoint" ,1 , 1 , font, paragraph, style);
  > new OfficeModelCreator().CreateTableCell(table1, 0, 1, "Word", 1 , 1 , font, paragraph, style);
  > new OfficeModelCreator().CreateTableCell(table1, 1, 0, "조금씩", 1 , 1 , font, paragraph, style);
  > new OfficeModelCreator().CreateTableCell(table1, 1, 1, "다르다", 1 , 1 , font, paragraph, style);
  > ```
  >
  > 2. 셀병합
  > ``` c#
  > var table2 = new OfficeModelCreator().CreateTable(400, 30, 300, 300, 2, 2);
  > officeModels.Add(table2);
  > //colspan 에 2를 주고 해당 셀을 만들지 않으면 병합됩니다. 
  > new OfficeModelCreator().CreateTableCell(table2, 0, 0, "PowerPoint", 1, 2, null, null, style2);
  > new OfficeModelCreator().CreateTableCell(table2, 1, 0, "조금씩", 1, 1, null, null, style2);
  > new OfficeModelCreator().CreateTableCell(table2, 1, 1, "다르다", 1, 1, null, null, style2);
  > ```
  >
  > 3. Col & Row 사이즈 지정 및 이미지 추가
  > ``` c#
  > List<int> columnWidth = new List<int>();
  > columnWidth.Add(100);
  > columnWidth.Add(500);
  > List<int> rowHeight = new List<int>();
  > rowHeight.Add(100);
  > rowHeight.Add(200);
  > 
  > //테이블 생성
  > var table3 = new OfficeModelCreator().CreateTable(50, 350, 600, 300, columnWidth, rowHeight);
  > officeModels.Add(table3);
  > 
  > new OfficeModelCreator().CreateTableCell(table3, 0, 0, "사이즈 지정", 1, 1, font, paragraph, style3);
  > new OfficeModelCreator().CreateTableCell(table3, 0, 1, "이미지 테이블내 삽입 워드만 가능", 1, 1, font, paragraph, style3);
  > new OfficeModelCreator().CreateTableCell(table3, 1, 0, "이미지", 1, 1, font, paragraph, style2);
  > 
  > //이미지를 테이블내에 삽입하는건 워드만 가능합니다. PPT는 해당 기능이 없이 그냥 지정한 위치에 이미지가 올라가게 됩니다. 
  > if(officeCreator.OfficeType == OfficeType.Word)
  > {
  >     string base64String = Utils.Instance.GetWebImage("https://img-prod-cms-rt-microsoft-com.akamaized.net/cms/api/am/imageFileData/RE1Mu3b?ver=5c31");
  >     //이미지 생성
  >     var image = new OfficeModelCreator().CreatePicture(0, 0, 100, 100, base64String);
  >     //테이블 내에 넣는 이미지라서 타입을 변경한다. 
  >     image.Type = Modules.Model.Type.TableImageCell;
  >     officeModels.Add(image);
  >     new OfficeModelCreator().CreateTableCell(table3, 1, 1, "", 1, 1, font, paragraph, style2, image.UID);
  > }
  > else
  > {
  >     new OfficeModelCreator().CreateTableCell(table3, 1, 0, "이 좌표 계산해서 이미지를 추가한다. ", 1, 1, font, paragraph, style2);
  > }
  > ```
<details><summary><h3>전체코드</h3></summary>    
  
``` C#
public string CreateSingleTableDocument()
{
    //생성할 문서를 지정한다.
    var officeCreator = new OfficeCreator(OfficeType.PowerPoint);
    //모델 목록을 만들고 테이블을 추가한다. 
    var officeModels = new List<OfficeModel>();

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
        TextDirection = TextDirection.Horizontal,
        LineSpacing = 0
    };
    //테이블 스타일을 설정합니다.[옵션]
    var style = new OfficeTableStyles()
    {
        UseShading = true,
        ShadingColor = "yellow"
    };
    var border = new Modules.Model.Component.TableDesignTab.Border()
    {
        Draw = true, 
        Color = "red",
        Weight = 1,
        Style = "solid"
    };
    style.Bottom = border;

    var style2 = new OfficeTableStyles()
    {
        UseShading = true,
        ShadingColor = "transparent"
    };
    var style3 = new OfficeTableStyles()
    {
        UseShading = true,
        ShadingColor = "green"
    };

    #region Case1
    var table1 = new OfficeModelCreator().CreateTable(50, 30, 300, 300, 2, 2);
    officeModels.Add(table1);

    new OfficeModelCreator().CreateTableCell(table1, 0, 0, "PowerPoint" ,1 , 1 , font, paragraph, style);
    new OfficeModelCreator().CreateTableCell(table1, 0, 1, "Word", 1 , 1 , font, paragraph, style);
    new OfficeModelCreator().CreateTableCell(table1, 1, 0, "조금씩", 1 , 1 , font, paragraph, style);
    new OfficeModelCreator().CreateTableCell(table1, 1, 1, "다르다", 1 , 1 , font, paragraph, style);
    #endregion

    #region Case2 셀 병합

    var table2 = new OfficeModelCreator().CreateTable(400, 30, 300, 300, 2, 2);
    officeModels.Add(table2);
    
    new OfficeModelCreator().CreateTableCell(table2, 0, 0, "PowerPoint", 1, 2, null, null, style2);
    new OfficeModelCreator().CreateTableCell(table2, 1, 0, "조금씩", 1, 1, null, null, style2);
    new OfficeModelCreator().CreateTableCell(table2, 1, 1, "다르다", 1, 1, null, null, style2);
    #endregion

    #region Case2 Col & Row 사이즈 지정 및 이미지 추가
    List<int> columnWidth = new List<int>();
    columnWidth.Add(100);
    columnWidth.Add(500);
    List<int> rowHeight = new List<int>();
    rowHeight.Add(100);
    rowHeight.Add(200);

    //테이블 생성
    var table3 = new OfficeModelCreator().CreateTable(50, 350, 600, 300, columnWidth, rowHeight);
    officeModels.Add(table3);
    
    new OfficeModelCreator().CreateTableCell(table3, 0, 0, "사이즈 지정", 1, 1, font, paragraph, style3);
    new OfficeModelCreator().CreateTableCell(table3, 0, 1, "이미지 테이블내 삽입 워드만 가능", 1, 1, font, paragraph, style3);
    new OfficeModelCreator().CreateTableCell(table3, 1, 0, "이미지", 1, 1, font, paragraph, style2);

    //이미지를 테이블내에 삽입하는건 워드만 가능합니다. PPT는 해당 기능이 없이 그냥 지정한 위치에 이미지가 올라가게 됩니다. 
    if(officeCreator.OfficeType == OfficeType.Word)
    {
        string base64String = Utils.Instance.GetWebImage("https://img-prod-cms-rt-microsoft-com.akamaized.net/cms/api/am/imageFileData/RE1Mu3b?ver=5c31");
        //이미지 생성
        var image = new OfficeModelCreator().CreatePicture(0, 0, 100, 100, base64String);
        //테이블 내에 넣는 이미지라서 타입을 변경한다. 
        image.Type = Modules.Model.Type.TableImageCell;
        officeModels.Add(image);
        new OfficeModelCreator().CreateTableCell(table3, 1, 1, "", 1, 1, font, paragraph, style2, image.UID);
    }
    else
    {
        new OfficeModelCreator().CreateTableCell(table3, 1, 0, "이 좌표 계산해서 이미지를 추가한다. ", 1, 1, font, paragraph, style2);
    }
    #endregion

    //변환
    officeCreator.ConvertPage(1, officeModels);

    return officeCreator.Save();
}
```
</details>

---
# 공통속성
- 각 컨트롤 여백
- 용지 사이즈
``` c#
 public string CreateCommonProperty()
 {
     //생성할 문서를 지정한다.
     //오피스에는 페이지별로 용지 사이즈 설정기능이 없습니다.
     var officeCreator = new OfficeCreator(OfficeType.Word, 600, 700);

     {   1번 페이지 생성
         //여백이 정상인지 확인하기 위해서 테두리 표시
         var style = new OfficeShapeStyle()
         {
             UseOutline = true
         };

         //텍스트 박스 생성
         var model = new OfficeModelCreator().CreateTextBox(50, 50, 100, 100, "1번페이지", null, null, style);
         model.Margin.Left = 20;
         model.Margin.Top = 50;
         //모델 목록을 만들고 생성한 컨트롤을 추가한다. 
         var officeModels = new List<OfficeModel>();
         officeModels.Add(model);

         //변환한다. 
         officeCreator.ConvertPage(1, officeModels);
     }
     {   2번 페이지 생성
         //텍스트 박스 생성
         var model = new OfficeModelCreator().CreateTextBox(50, 50, 100, 100, "2번 페이지");
         //모델 목록을 만들고 생성한 컨트롤을 추가한다. 
         var officeModels = new List<OfficeModel>();
         officeModels.Add(model);

         //변환한다. 
         officeCreator.ConvertPage(2, officeModels);
     }
     return officeCreator.Save();

 }
```
  


### 기타

- UI에서 글자 간격 상세 옵션 화면

|Word|Powerpoint|
|:---:|:---:|
|![image](https://github.com/jyseok85/SimpleOfficeCreator.Stardard/assets/48501866/d2e28df7-0975-4f58-98de-95f86d03f39b)|![image](https://github.com/jyseok85/SimpleOfficeCreator.Stardard/assets/48501866/a02d33dd-8712-4134-8a54-edffffef783f)|

- Word 테이블 셀에는 글자 자동맞춤 기능이 있습니다.  
  위에서 설명하지는 않았지만 단락설정의 TableCellFitTextForWord 값으로 설정가능합니다.

||
|:---:|
|![image](https://github.com/jyseok85/SimpleOfficeCreator.Stardard/assets/48501866/f747cdd5-b653-4900-b460-bf8a3e1fcd29)|

