using SimpleOfficeCreator.Stardard.Modules;
using SimpleOfficeCreator.Stardard.Modules.GeneratedCode;
using SimpleOfficeCreator.Stardard.Modules.Model;
using System;
using System.Collections.Generic;
using System.IO;

namespace SimpleOfficeCreator.Stardard
{
    public enum OfficeType
    {
        None,
        Excel,
        PowerPoint,
        Word
    }

    public class OfficeCreator
    {
        // EMU(English Metric Units) 96ppi 이미지의 변환 계수는 9525이고, 72ppi 이미지의 경우 12700, 300ppi 이미지의 경우 3048
        const int EMU96PPI = 9525;
        //const string IMAGE_NOIMAGE = @"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";
        PowerPoint powerPoint = null;
        Word word = null;
        MemoryStream memoryStream = new System.IO.MemoryStream();


        OfficeType OfficeType { get; set; }

        public void Initialize(OfficeType type)
        {
            this.OfficeType = type;
            this.memoryStream = new MemoryStream();

            Common.Instance.UniqueId.Clear();
            Common.Instance.UniqueId.Add(1);
            Common.Instance.EMUPPI = EMU96PPI;

            switch (OfficeType)
            {
                case OfficeType.PowerPoint:
                    powerPoint = new PowerPoint(memoryStream);
                    //to do:PPT 용지사이즈 설정
                    powerPoint.Initialize(794, 1123);
                    break;
                case OfficeType.Word:
                    word = new Word(memoryStream);
                    word.Initialize(800, 1123);
                    break;
                case OfficeType.Excel:
                    break;
            }
        }

        public void Convert(List<OfficeModel> models)
        {
            //오피스 오브젝트로 만들고 넘긴다. 
            //즉 이 클래스는 기존 모듈과 분리된다. 
            powerPoint.ConvertPerPage(1, models);
        }

        public void ConvertPage(int page, List<OfficeModel> models)
        {
            switch (OfficeType)
            {
                case OfficeType.PowerPoint:
                    powerPoint.ConvertPerPage(page, models);
                    break;
                case OfficeType.Word:
                    word.ConvertPerPage(page, models);
                    break;
                case OfficeType.Excel:
                    break;
            }
        }

        /// <summary>
        /// 파일을 저장합니다.
        /// </summary>
        /// <param name="filePath"></param>
        public void Save(string filePath)
        {
            SaveOfficeDocument();

            if (filePath == string.Empty)
            {
                switch (OfficeType)
                {
                    case OfficeType.PowerPoint:
                        filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads", OfficeType.ToString() + ".pptx");
                        break;
                    case OfficeType.Word:
                        filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads", OfficeType.ToString() + ".docx");
                        break;
                    case OfficeType.Excel:
                        filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads", OfficeType.ToString() + ".xlsx");
                        break;
                }
            }

            File.WriteAllBytes(filePath, this.memoryStream.ToArray());
            this.memoryStream.Dispose();
        }

        /// <summary>
        /// 현재 생성한 문서를 가져옵니다.
        /// </summary>
        /// <returns></returns>
        public byte[] GetByteArray()
        {
            SaveOfficeDocument();

            this.memoryStream.Seek(0, SeekOrigin.Begin);
            var result = this.memoryStream.ToArray();
            this.memoryStream.Dispose();
            return result;
        }

        private void SaveOfficeDocument()
        {
            switch (OfficeType)
            {
                case OfficeType.PowerPoint:
                    powerPoint.Save();
                    break;
                case OfficeType.Word:
                    word.Save();
                    break;
            }
        }
    }
}
