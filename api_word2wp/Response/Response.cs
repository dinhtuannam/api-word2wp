using System.ComponentModel;

namespace api_word2wp.Response
{
    public class ResponseResult<T> where T : class
    {
        public ResponseResult() { }

        public ResponseResult(RetCodeEnum retCode, string retText, T data)
        {
            this.RetCode = retCode;
            this.RetText = retText;
            switch (retCode)
            {
                case RetCodeEnum.ApiNoDelete:
                    this.RetText = "Cần phải xóa cấp con trước khi xóa.";
                    break;
                case RetCodeEnum.ApiNotRole:
                    this.RetText = "Bạn không có quyền.";
                    break;
            }
            this.Data = data;
        }

        public RetCodeEnum RetCode { get; set; }
        public string RetText { get; set; }
        public T Data { get; set; }
    }

    public enum RetCodeEnum
    {
        [Description("OK")]
        Ok = 0,
        [Description("Api Error")]
        ApiError = 1,
        [Description("Not Exists")]
        ResultNotExists = 2,
        [Description("Parammeters Invalid")]
        ParammetersInvalid = 3,
        [Description("Parammeters Not Found")]
        ParammetersNotFound = 4,
        [Description("Not delete")]
        ApiNoDelete = 5,
        [Description("Not Role")]
        ApiNotRole = 6
    }
}
