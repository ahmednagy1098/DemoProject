using System;
using System.Data;
using System.IO;
using System.Text;

public class ReportHelper
{
    public static MemoryStream GenerateDynamicRDLC(DataTable dt, string datasetName = "MyDataSet")
    {
        StringBuilder sb = new StringBuilder();

        sb.AppendLine(@"<?xml version=""1.0"" encoding=""utf-8""?>");
        sb.AppendLine(@"<Report xmlns=""http://schemas.microsoft.com/sqlserver/reporting/2008/01/reportdefinition"" xmlns:rd=""http://schemas.microsoft.com/SQLServer/reportdesigner"">");
        sb.AppendLine(@"  <DataSources>");
        sb.AppendLine(@"    <DataSource Name=""DummyDataSource"">");
        sb.AppendLine(@"      <ConnectionProperties>");
        sb.AppendLine(@"        <DataProvider>System.Data.DataSet</DataProvider>");
        sb.AppendLine(@"        <ConnectString>/* Local Connection */</ConnectString>");
        sb.AppendLine(@"      </ConnectionProperties>");
        sb.AppendLine(@"      <rd:DataSourceID>" + Guid.NewGuid().ToString() + @"</rd:DataSourceID>");
        sb.AppendLine(@"    </DataSource>");
        sb.AppendLine(@"  </DataSources>");

        // 🟢 تعريف الـ Dataset
        sb.AppendLine(@"  <DataSets>");
        sb.AppendLine(@"    <DataSet Name=""" + datasetName + @""">");
        sb.AppendLine(@"      <Fields>");
        foreach (DataColumn col in dt.Columns)
        {
            sb.AppendLine(@"        <Field Name=""" + col.ColumnName + @""">");
            sb.AppendLine(@"          <DataField>" + col.ColumnName + @"</DataField>");
            sb.AppendLine(@"        </Field>");
        }
        sb.AppendLine(@"      </Fields>");
        sb.AppendLine(@"      <Query>");
        sb.AppendLine(@"        <DataSourceName>DummyDataSource</DataSourceName>");
        sb.AppendLine(@"        <CommandText>/* Local Query */</CommandText>");
        sb.AppendLine(@"      </Query>");
        sb.AppendLine(@"    </DataSet>");
        sb.AppendLine(@"  </DataSets>");

        // 🟢 Body (Table/Tablix)
        sb.AppendLine(@"  <Body>");
        sb.AppendLine(@"    <ReportItems>");
        sb.AppendLine(@"      <Tablix Name=""Tablix1"">");
        sb.AppendLine(@"        <TablixBody>");
        sb.AppendLine(@"          <TablixColumns>");
        foreach (DataColumn col in dt.Columns)
        {
            sb.AppendLine(@"            <TablixColumn>");
            sb.AppendLine(@"              <Width>2cm</Width>");
            sb.AppendLine(@"            </TablixColumn>");
        }
        sb.AppendLine(@"          </TablixColumns>");

        // 🟢 صف واحد للرأس + صف واحد للبيانات
        sb.AppendLine(@"          <TablixRows>");
        // Header Row
        sb.AppendLine(@"            <TablixRow>");
        sb.AppendLine(@"              <Height>0.6cm</Height>");
        sb.AppendLine(@"              <TablixCells>");
        foreach (DataColumn col in dt.Columns)
        {
            sb.AppendLine(@"                <TablixCell>");
            sb.AppendLine(@"                  <CellContents>");
            sb.AppendLine(@"                    <Textbox Name=""Header_" + col.ColumnName + @""">");
            sb.AppendLine(@"                      <Paragraphs>");
            sb.AppendLine(@"                        <Paragraph>");
            sb.AppendLine(@"                          <TextRuns>");
            sb.AppendLine(@"                            <TextRun>");
            sb.AppendLine(@"                              <Value>" + col.ColumnName + @"</Value>");
            sb.AppendLine(@"                              <Style>");
            sb.AppendLine(@"                                <FontWeight>Bold</FontWeight>");
            sb.AppendLine(@"                              </Style>");
            sb.AppendLine(@"                            </TextRun>");
            sb.AppendLine(@"                          </TextRuns>");
            sb.AppendLine(@"                        </Paragraph>");
            sb.AppendLine(@"                      </Paragraphs>");
            sb.AppendLine(@"                      <Style />");
            sb.AppendLine(@"                    </Textbox>");
            sb.AppendLine(@"                  </CellContents>");
            sb.AppendLine(@"                </TablixCell>");
        }
        sb.AppendLine(@"              </TablixCells>");
        sb.AppendLine(@"            </TablixRow>");

        // Data Row
        sb.AppendLine(@"            <TablixRow>");
        sb.AppendLine(@"              <Height>0.6cm</Height>");
        sb.AppendLine(@"              <TablixCells>");
        foreach (DataColumn col in dt.Columns)
        {
            sb.AppendLine(@"                <TablixCell>");
            sb.AppendLine(@"                  <CellContents>");
            sb.AppendLine(@"                    <Textbox Name=""Data_" + col.ColumnName + @""">");
            sb.AppendLine(@"                      <Paragraphs>");
            sb.AppendLine(@"                        <Paragraph>");
            sb.AppendLine(@"                          <TextRuns>");
            sb.AppendLine(@"                            <TextRun>");
            sb.AppendLine(@"                              <Value>=Fields!" + col.ColumnName + @".Value</Value>");
            sb.AppendLine(@"                            </TextRun>");
            sb.AppendLine(@"                          </TextRuns>");
            sb.AppendLine(@"                        </Paragraph>");
            sb.AppendLine(@"                      </Paragraphs>");
            sb.AppendLine(@"                      <Style />");
            sb.AppendLine(@"                    </Textbox>");
            sb.AppendLine(@"                  </CellContents>");
            sb.AppendLine(@"                </TablixCell>");
        }
        sb.AppendLine(@"              </TablixCells>");
        sb.AppendLine(@"            </TablixRow>");

        sb.AppendLine(@"          </TablixRows>");
        sb.AppendLine(@"        </TablixBody>");
        sb.AppendLine(@"        <DataSetName>" + datasetName + @"</DataSetName>");
        sb.AppendLine(@"      </Tablix>");
        sb.AppendLine(@"    </ReportItems>");
        sb.AppendLine(@"    <Height>5cm</Height>");
        sb.AppendLine(@"  </Body>");

        sb.AppendLine(@"  <Width>20cm</Width>");
        sb.AppendLine(@"  <Page>");
        sb.AppendLine(@"    <PageHeight>29.7cm</PageHeight>");
        sb.AppendLine(@"    <PageWidth>21cm</PageWidth>");
        sb.AppendLine(@"    <LeftMargin>2cm</LeftMargin>");
        sb.AppendLine(@"    <RightMargin>2cm</RightMargin>");
        sb.AppendLine(@"    <TopMargin>2cm</TopMargin>");
        sb.AppendLine(@"    <BottomMargin>2cm</BottomMargin>");
        sb.AppendLine(@"  </Page>");

        sb.AppendLine(@"</Report>");

        return new MemoryStream(Encoding.UTF8.GetBytes(sb.ToString()));
    }
}
