using ExcelTools.Abstraction;
using ExcelTools.Cleaner;
using ExcelTools.ColumnSplitter;
using ExcelTools.DuplicateRemover;
using ExcelTools.Merger;
using ExcelTools.Splitter;
using System.Reflection;

namespace ExcelToolsTests
{
    public class UnitTest1: ExcelOptionsBase
    {
        string currentDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

        [Fact]
        public void Cleaner()
        {
            var inputFilePath = Path.Combine(currentDirectory, "TestFiles", "Cleaner", "cleaner_input.xlsx");
            var outputFilePath = Path.Combine(currentDirectory, "TestFiles", "Cleaner", "new_cleaner_input.xlsx");

            var cleaner = new Cleaner();

            var result = cleaner.Process(new CleanOptions
            {
                FilePath = inputFilePath,
                ResultFilePath = outputFilePath
            });

            Assert.Equal(ResultCode.Success, result.Code);
        }

        [Fact]
        public void ColumnSplitter()
        {
            var inputFilePath = Path.Combine(currentDirectory, "TestFiles", "ColumnSplitter", "column_split_input.xlsx");
            var outputFilePath = Path.Combine(currentDirectory, "TestFiles", "ColumnSplitter", "new_column_split_input.xlsx");

            var columnSplitter = new ColumnSplitter();

            var Result = columnSplitter.Process(new ColumnSplitterOptions()
            {
                FilePath = inputFilePath,
                ResultFilePath = outputFilePath,
                ColumnName = "D",
                SheetNumber = 1,
                SkipHeaderRows = 1,
                SplitSymbols = " "
            });

            Assert.Equal(ResultCode.Success, Result.Code);
        }

        [Fact]
        public void DuplicateRemover()
        {
            var inputFilePath = Path.Combine(currentDirectory, "TestFiles", "DuplicateRemover", "duplicateRemover_input.xlsx");
            var outputFilePath = Path.Combine(currentDirectory, "TestFiles", "DuplicateRemover", "new_duplicateRemover_input.xlsx");

            var duplicateRemover = new DuplicateRemover();

            var Result = duplicateRemover.Process(new DuplicateRemoverOptions()
            {
                FilePath = inputFilePath,
                ResultFilePath = outputFilePath,
            });

            Assert.Equal(ResultCode.Success, Result.Code);
        }

        [Fact]
        public void Merger()
        {
            var inputFilePath1 = Path.Combine(currentDirectory, "TestFiles", "Merger", "merge_input_1.xlsx");
            var inputFilePath2 = Path.Combine(currentDirectory, "TestFiles", "Merger", "merge_input_2.xlsx");
            var inputFilePath3 = Path.Combine(currentDirectory, "TestFiles", "Merger", "merge_input_3.xlsx");

            var outputFilePath = Path.Combine(currentDirectory, "TestFiles", "Merger", "new_merge_input.xlsx");

            var merger = new Merger();

            var Result = merger.Process(new MergerOptions()
            {
                MergeFilePaths = new []{ inputFilePath1, inputFilePath2, inputFilePath3 },
                ResultFilePath = outputFilePath
            });

            Assert.Equal(ResultCode.Success, Result.Code);
        }

        [Fact]
        public void Splitter()
        {
            var inputFilePath = Path.Combine(currentDirectory, "TestFiles", "Splitter", "split_input.xlsx");
            var outputFilePath = Path.Combine(currentDirectory, "TestFiles", "Splitter", "new_split_input_{0}.xlsx");

            var splitter = new Splitter();

            var Result = splitter.Process(new SplitterOptions()
            {
                FilePath = inputFilePath,
                ResultFilePath = outputFilePath, 
                AddHeaderRows = 1, 
                ResultsCount = 20, 
                SplitMode = SplitterOptions.SplitType.SplitByRows
            });

            Assert.Equal(ResultCode.Success, Result.Code);
        }
    }
}