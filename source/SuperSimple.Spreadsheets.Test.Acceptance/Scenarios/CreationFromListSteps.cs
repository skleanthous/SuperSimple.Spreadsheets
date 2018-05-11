using SuperSimple.Spreadsheets; 
using System;
using System.IO;
using TechTalk.SpecFlow;
using FluentAssertions;
using System.Linq;

namespace SuperSimple.Spreadsheets.Test.Acceptance.Scenarios
{
    [Binding]
    public class CreationFromListSteps
    {
        private const string DATA_CONTEXT_KEY = "DataSaveFromListTest";
        private const string STREAM_CONTEXT_KEY = "SavedStreamFromListTest";

        private class Data
        {
            public int ID { get; set; }
            public string Title { get; set; }
            public string Author { get; set; }
        }

        private Data[] ToStore
        {
            get { return (Data[])ScenarioContext.Current[DATA_CONTEXT_KEY]; }
            set { ScenarioContext.Current[DATA_CONTEXT_KEY] = value; }
        }

        private MemoryStream SpreadsheetStream
        {
            get { return (MemoryStream)ScenarioContext.Current[STREAM_CONTEXT_KEY]; }
            set { ScenarioContext.Current[STREAM_CONTEXT_KEY] = value; }
        }

        [Given(@"a list containing (.*) items of a specific type")]
        public void GivenAListContainingItemsOfASpecificType(int p0)
        {
            var datas = new Data[p0];

            for(int i =0;i<p0;i++)
            {
                datas[i] = new Data()
                {
                    ID = i,
                    Author = String.Format("{0}{0}{0}", (char)('A' + i)),
                    Title = String.Format("{0}{0}{0}", (char)('a' + i)),
                };
            }

            ToStore = datas;
        }

        [When(@"I call SaveToStream with the available data")]
        public void WhenICreateASaverWithTheAvailableData()
        {
            SpreadsheetStream = new MemoryStream();

            ExcelSaver.Save(ToStore, SpreadsheetStream);
        }

        [Then(@"the result should be a file that can be opened")]
        public void ThenTheResultShouldBeAFileThatCanBeOpened()
        {
            SpreadsheetStream.Seek(0, SeekOrigin.Begin);

            ExcelLoader.LoadReadOnlyFromStream(SpreadsheetStream).ReadRows();
        }

        [Then(@"it should contain (.*) rows \((.*) for items and one for header\)")]
        public void ThenItShouldContainRowsForItemsAndOneForHeader(int p0, int p1)
        {
            SpreadsheetStream.Seek(0, SeekOrigin.Begin);

            ExcelLoader.LoadReadOnlyFromStream(SpreadsheetStream)
                .ReadRows()
                .Count.Should().Be(p0);
        }

        [Then(@"each row should correspond to the data in the list")]
        public void ThenEachRowShouldCorrespondToTheDataInTheList()
        {
            SpreadsheetStream.Seek(0, SeekOrigin.Begin);

            var rows = ExcelLoader.LoadReadOnlyFromStream(SpreadsheetStream)
                            .ReadRows()
                            .ToArray();

            //The first row is the titles so we kip it
            for(int i =1;i<rows.Length;i++)
            {
                rows[i].Any(x => x.ValueType == typeof(long) && x.Value == ToStore[i - 1].ID).Should().BeTrue();
                rows[i].Any(x => x.ValueType == typeof(string) && x.Value == ToStore[i - 1].Title).Should().BeTrue();
                rows[i].Any(x => x.ValueType == typeof(String) && x.Value == ToStore[i - 1].Author).Should().BeTrue();
            }
        }
    }
}
