using System.ComponentModel.DataAnnotations;

namespace DictionaryManagement_Models.IntDBModels
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    public class CheckLeadingAndTrailingSpaces : ValidationAttribute
    {
        public override bool IsValid(object? value)
        {
            var isValid = true;
            if (value != null)
            {
                var inputValue = value as string;
                if (inputValue != null)
                {
                    if (inputValue.Length > 0)
                    {
                        if (Char.IsWhiteSpace(inputValue[0]))
                        {
                            isValid = false;
                        }
                    }
                    if (inputValue.Length >= 2)
                        if (Char.IsWhiteSpace(inputValue[inputValue.Length - 1]))
                        {
                            isValid = false;
                        }
                }
            }
            if (isValid != true)
                this.ErrorMessage = "Поле начинается или заканчивается пробелом или другим пробельным символом";
            return isValid;
        }
    }
}
