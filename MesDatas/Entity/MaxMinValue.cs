namespace MesDatas.Entity
{
    public class MaxMinValue
    {
        public string BoardName { get; set; }       // 测试项
        public string StandardCode { get; set; }    // 标准值
        public string MaxBoardCode { get; set; }    // 上限值
        public string MinBoardCode { get; set; }    // 下限值
        public string BoardCode { get; set; }       // 实际值
        public string Result { get; set; }          // 测试结果
    }
}
