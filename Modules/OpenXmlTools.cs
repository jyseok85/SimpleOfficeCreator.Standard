﻿using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Security.Cryptography;
using System.Text;

namespace OpenXMLTools
{
    public static class Tools
    {
        /// <summary>
        /// Word파일에 암호를 설정합니다.
        /// 파일열때 암호는 다른 외부 모듈을 사용해서 설정해야합니다. 
        /// 외부모듈 사용시 (FileOpen 프로세스가 들어가야 하기 때문에 브라우저상에서는 불가능합니다.)
        /// </summary>
        /// <returns></returns>
        public static string ApplyDocumentProtection(WordprocessingDocument document, string password = "")
        {
            if (password == string.Empty)
                return string.Empty;

            // generate a 128-bit salt using a secure PRNG
            byte[] salt = new byte[128 / 8];
            using (RandomNumberGenerator rng = RandomNumberGenerator.Create())
            {
                rng.GetBytes(salt);
            }
            string pasalt = Convert.ToBase64String(salt);

            //word 용 해시 생성
            //https://social.msdn.microsoft.com/Forums/vstudio/en-US/63588f50-354f-43ba-b080-e0e6c51a0fb5/hash-and-saltdocumentprotection?forum=clr
            string hashed = GetHashString(password, pasalt, 100000);

            DocumentProtection documentProtection = new DocumentProtection
            {
                Edit = DocumentProtectionValues.ReadOnly,
                Enforcement = new OnOffValue(true),
                CryptographicAlgorithmClass = CryptAlgorithmClassValues.Hash,
                CryptographicProviderType = CryptProviderValues.RsaAdvancedEncryptionStandard,
                CryptographicAlgorithmType = CryptAlgorithmValues.TypeAny,
                CryptographicAlgorithmSid = 4, // SHA1
                CryptographicSpinCount = 100000,
                Hash = new Base64BinaryValue() { Value = hashed },
                Salt = new Base64BinaryValue() { Value = pasalt }
            };
            document.MainDocumentPart.DocumentSettingsPart.Settings.AppendChild(documentProtection);
            document.MainDocumentPart.DocumentSettingsPart.Settings.Save();
            return password;
        }



        public static string GetHashString(string strPassword, string Salt, int cryptSpinCount)
        {
            // Generate the Salt
            byte[] arrSalt = Convert.FromBase64String(Salt);

            //Array to hold Key Values
            byte[] generatedKey = new byte[4];

            //Maximum length of the password is 15 chars.
            int intMaxPasswordLength = 15;


            if (!String.IsNullOrEmpty(strPassword))
            {
                // Truncate the password to 15 characters
                strPassword = strPassword.Substring(0, Math.Min(strPassword.Length, intMaxPasswordLength));

                // Construct a new NULL-terminated string consisting of single-byte characters:
                //  -- > Get the single-byte values by iterating through the Unicode characters of the truncated Password.
                //   --> For each character, if the low byte is not equal to 0, take it. Otherwise, take the high byte.

                byte[] arrByteChars = new byte[strPassword.Length];

                for (int intLoop = 0; intLoop < strPassword.Length; intLoop++)
                {
                    int intTemp = Convert.ToInt32(strPassword[intLoop]);
                    arrByteChars[intLoop] = Convert.ToByte(intTemp & 0x00FF);
                    if (arrByteChars[intLoop] == 0)
                        arrByteChars[intLoop] = Convert.ToByte((intTemp & 0xFF00) >> 8);
                }

                // Compute the high-order word of the new key:

                // --> Initialize from the initial code array (see below), depending on the strPassword’s length. 
                int intHighOrderWord = InitialCodeArray[arrByteChars.Length - 1];

                // --> For each character in the strPassword:
                //      --> For every bit in the character, starting with the least significant and progressing to (but excluding) 
                //          the most significant, if the bit is set, XOR the key’s high-order word with the corresponding word from 
                //          the Encryption Matrix

                for (int intLoop = 0; intLoop < arrByteChars.Length; intLoop++)
                {
                    int tmp = intMaxPasswordLength - arrByteChars.Length + intLoop;
                    for (int intBit = 0; intBit < 7; intBit++)
                    {
                        if ((arrByteChars[intLoop] & (0x0001 << intBit)) != 0)
                        {
                            intHighOrderWord ^= EncryptionMatrix[tmp, intBit];
                        }
                    }
                }

                // Compute the low-order word of the new key:

                // Initialize with 0
                int intLowOrderWord = 0;

                // For each character in the strPassword, going backwards
                for (int intLoopChar = arrByteChars.Length - 1; intLoopChar >= 0; intLoopChar--)
                {
                    // low-order word = (((low-order word SHR 14) AND 0x0001) OR (low-order word SHL 1) AND 0x7FFF)) XOR character
                    intLowOrderWord = (((intLowOrderWord >> 14) & 0x0001) | ((intLowOrderWord << 1) & 0x7FFF)) ^ arrByteChars[intLoopChar];
                }

                // Lastly,low-order word = (((low-order word SHR 14) AND 0x0001) OR (low-order word SHL 1) AND 0x7FFF)) XOR strPassword length XOR 0xCE4B.
                intLowOrderWord = (((intLowOrderWord >> 14) & 0x0001) | ((intLowOrderWord << 1) & 0x7FFF)) ^ arrByteChars.Length ^ 0xCE4B;

                // Combine the Low and High Order Word
                int intCombinedkey = (intHighOrderWord << 16) + intLowOrderWord;

                // The byte order of the result shall be reversed [Example: 0x64CEED7E becomes 7EEDCE64. end example],
                // and that value shall be hashed as defined by the attribute values.

                for (int intTemp = 0; intTemp < 4; intTemp++)
                {
                    generatedKey[intTemp] = Convert.ToByte(((uint)(intCombinedkey & (0x000000FF << (intTemp * 8)))) >> (intTemp * 8));
                }
            }

            // Implementation Notes List:
            // --> In this third stage, the reversed byte order legacy hash from the second stage shall be converted to Unicode hex 
            // --> string representation 
            StringBuilder sb = new StringBuilder();
            for (int intTemp = 0; intTemp < 4; intTemp++)
            {
                sb.Append(Convert.ToString(generatedKey[intTemp], 16));
            }

            generatedKey = Encoding.Unicode.GetBytes(sb.ToString().ToUpper());

            // Implementation Notes List:
            //Word appends the binary form of the salt attribute and not the base64 string representation when hashing
            // Before calculating the initial hash, you are supposed to prepend (not append) the salt to the key
            byte[] tmpArray1 = generatedKey;
            byte[] tmpArray2 = arrSalt;
            byte[] tempKey = new byte[tmpArray1.Length + tmpArray2.Length];
            Buffer.BlockCopy(tmpArray2, 0, tempKey, 0, tmpArray2.Length);
            Buffer.BlockCopy(tmpArray1, 0, tempKey, tmpArray2.Length, tmpArray1.Length);
            generatedKey = tempKey;


            // Iterations specifies the number of times the hashing function shall be iteratively run (using each
            // iteration's result as the input for the next iteration).
            int iterations = cryptSpinCount;

            // Implementation Notes List:
            //Word requires that the initial hash of the password with the salt not be considered in the count.
            //    The initial hash of salt + key is not included in the iteration count.
            HashAlgorithm sha1 = new SHA1Managed();
            generatedKey = sha1.ComputeHash(generatedKey);
            byte[] iterator = new byte[4];
            for (int intTmp = 0; intTmp < iterations; intTmp++)
            {

                //When iterating on the hash, you are supposed to append the current iteration number.
                iterator[0] = Convert.ToByte((intTmp & 0x000000FF) >> 0);
                iterator[1] = Convert.ToByte((intTmp & 0x0000FF00) >> 8);
                iterator[2] = Convert.ToByte((intTmp & 0x00FF0000) >> 16);
                iterator[3] = Convert.ToByte((intTmp & 0xFF000000) >> 24);

                generatedKey = ConcatByteArrays(iterator, generatedKey);
                generatedKey = sha1.ComputeHash(generatedKey);
            }

            // Apply the element
            return Convert.ToBase64String(generatedKey);

        }
        static readonly int[] InitialCodeArray = { 0xE1F0, 0x1D0F, 0xCC9C, 0x84C0, 0x110C, 0x0E10, 0xF1CE, 0x313E, 0x1872, 0xE139, 0xD40F, 0x84F9, 0x280C, 0xA96A, 0x4EC3 };
        static readonly int[,] EncryptionMatrix = new int[15, 7]
        {
            
            /* char 1  */ {0xAEFC, 0x4DD9, 0x9BB2, 0x2745, 0x4E8A, 0x9D14, 0x2A09},
            /* char 2  */ {0x7B61, 0xF6C2, 0xFDA5, 0xEB6B, 0xC6F7, 0x9DCF, 0x2BBF},
            /* char 3  */ {0x4563, 0x8AC6, 0x05AD, 0x0B5A, 0x16B4, 0x2D68, 0x5AD0},
            /* char 4  */ {0x0375, 0x06EA, 0x0DD4, 0x1BA8, 0x3750, 0x6EA0, 0xDD40},
            /* char 5  */ {0xD849, 0xA0B3, 0x5147, 0xA28E, 0x553D, 0xAA7A, 0x44D5},
            /* char 6  */ {0x6F45, 0xDE8A, 0xAD35, 0x4A4B, 0x9496, 0x390D, 0x721A},
            /* char 7  */ {0xEB23, 0xC667, 0x9CEF, 0x29FF, 0x53FE, 0xA7FC, 0x5FD9},
            /* char 8  */ {0x47D3, 0x8FA6, 0x0F6D, 0x1EDA, 0x3DB4, 0x7B68, 0xF6D0},
            /* char 9  */ {0xB861, 0x60E3, 0xC1C6, 0x93AD, 0x377B, 0x6EF6, 0xDDEC},
            /* char 10 */ {0x45A0, 0x8B40, 0x06A1, 0x0D42, 0x1A84, 0x3508, 0x6A10},
            /* char 11 */ {0xAA51, 0x4483, 0x8906, 0x022D, 0x045A, 0x08B4, 0x1168},
            /* char 12 */ {0x76B4, 0xED68, 0xCAF1, 0x85C3, 0x1BA7, 0x374E, 0x6E9C},
            /* char 13 */ {0x3730, 0x6E60, 0xDCC0, 0xA9A1, 0x4363, 0x86C6, 0x1DAD},
            /* char 14 */ {0x3331, 0x6662, 0xCCC4, 0x89A9, 0x0373, 0x06E6, 0x0DCC},
            /* char 15 */ {0x1021, 0x2042, 0x4084, 0x8108, 0x1231, 0x2462, 0x48C4}
      };

        private static byte[] ConcatByteArrays(byte[] array1, byte[] array2)
        {
            byte[] result = new byte[array1.Length + array2.Length];
            Buffer.BlockCopy(array2, 0, result, 0, array2.Length);
            Buffer.BlockCopy(array1, 0, result, array2.Length, array1.Length);
            return result;
        }
    }
}
