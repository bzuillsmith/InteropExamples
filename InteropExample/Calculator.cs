using System.Runtime.InteropServices;

namespace InteropExample
{
    [ComVisible(true),
        Guid("475B6C20-B1F0-49B3-95EF-B4D18BE9084E")]
    public interface ICalculator
    {
        int Add(int i1, int i2);
    }

    [ComVisible(true),
        ClassInterface(ClassInterfaceType.None),
        Guid("33D6E6F7-C20A-4428-987B-1CC34B32616E")]
    public class Calculator : ICalculator
    {
        public int Add(int i1, int i2)
        {
            return i1 + i2;
        }
    }
}
