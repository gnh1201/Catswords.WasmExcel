namespace Catswords.WasmExcel.Model
{
    public class Cells
    {
        public Sheets Parent { get; set; }

        public Cells(Sheets parent)
        {
            Parent = parent;
        }
    }
}
