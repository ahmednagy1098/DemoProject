using System;
using System.Data;
using System.IO;
using System.Text;
using System.Collections.Generic;

namespace DemoProject
{
    public class ReportHelperEnhanced
    {
        public static MemoryStream GenerateDynamicRDLC(
            DataTable dt,
            Dictionary<string, string> columnMap,
            string datasetName = "MyDataSet",
            string report_title = ""
            )
        {
            StringBuilder sb = new StringBuilder();

            sb.AppendLine(@"<?xml version=""1.0"" encoding=""utf-8""?>");
            sb.AppendLine(@"<Report xmlns=""http://schemas.microsoft.com/sqlserver/reporting/2008/01/reportdefinition"" xmlns:rd=""http://schemas.microsoft.com/SQLServer/reportdesigner"">");

            // ---------------- DataSource ----------------
            sb.AppendLine(@"  <DataSources>");
            sb.AppendLine(@"    <DataSource Name=""DummyDataSource"">");
            sb.AppendLine(@"      <ConnectionProperties>");
            sb.AppendLine(@"        <DataProvider>System.Data.DataSet</DataProvider>");
            sb.AppendLine(@"        <ConnectString>/* Local Connection */</ConnectString>");
            sb.AppendLine(@"      </ConnectionProperties>");
            sb.AppendLine(@"      <rd:DataSourceID>" + Guid.NewGuid().ToString() + @"</rd:DataSourceID>");
            sb.AppendLine(@"    </DataSource>");
            sb.AppendLine(@"  </DataSources>");

            // ---------------- DataSet ----------------
            sb.AppendLine(@"  <DataSets>");
            sb.AppendLine(@"    <DataSet Name=""" + datasetName + @""">");
            sb.AppendLine(@"      <Fields>");
            foreach (DataColumn col in dt.Columns)
            {
                string colName = col.ColumnName;
                string headerText = GetHeaderText(columnMap, colName);
                sb.AppendLine(@"        <Field Name=""" + colName + @""">");
                sb.AppendLine(@"          <DataField>" + colName + @"</DataField>");
                sb.AppendLine(@"        </Field>");
            }
            sb.AppendLine(@"      </Fields>");
            sb.AppendLine(@"      <Query>");
            sb.AppendLine(@"        <DataSourceName>DummyDataSource</DataSourceName>");
            sb.AppendLine(@"        <CommandText>/* Local Query */</CommandText>");
            sb.AppendLine(@"      </Query>");
            sb.AppendLine(@"    </DataSet>");
            sb.AppendLine(@"  </DataSets>");

            // ---------------- Body ----------------
            sb.AppendLine(@"  <Body>");
            sb.AppendLine(@"    <ReportItems>");

            // ===== Report Title =====
            sb.AppendLine(@"      <Textbox Name=""ReportTitle"">");
            sb.AppendLine(@"        <Top>0.2cm</Top>");

            // center textbox on A4 page (21cm wide, 1cm margins → ~19cm usable)
            // if textbox is 12cm, left offset = (19 - 12)/2 = 3.5cm
            sb.AppendLine(@"        <Left>3.5cm</Left>");
            sb.AppendLine(@"        <Height>1cm</Height>");
            sb.AppendLine(@"        <Width>12cm</Width>");

            sb.AppendLine(@"        <Paragraphs>");
            sb.AppendLine(@"          <Paragraph>");
            sb.AppendLine(@"            <TextRuns>");
            sb.AppendLine(@"              <TextRun>");
            sb.AppendLine($@"                <Value>{System.Security.SecurityElement.Escape(report_title)}</Value>");
            sb.AppendLine(@"                <Style>");
            sb.AppendLine(@"                  <FontSize>16pt</FontSize>");
            sb.AppendLine(@"                  <FontWeight>Bold</FontWeight>");
            sb.AppendLine(@"                  <Direction>RTL</Direction>");
            sb.AppendLine(@"                  <TextAlign>Center</TextAlign>");
            sb.AppendLine(@"                </Style>");
            sb.AppendLine(@"              </TextRun>");
            sb.AppendLine(@"            </TextRuns>");
            sb.AppendLine(@"          </Paragraph>");
            sb.AppendLine(@"        </Paragraphs>");
            sb.AppendLine(@"      </Textbox>");
            // ===== Execution Time =====
            sb.AppendLine(@"      <Textbox Name=""ExecutionTime"">");
            sb.AppendLine(@"        <Top>1.5cm</Top>");
            sb.AppendLine(@"        <Left>0.5cm</Left>");
            sb.AppendLine(@"        <Height>0.7cm</Height>");
            sb.AppendLine(@"        <Width>20cm</Width>");
            sb.AppendLine(@"        <Paragraphs>");
            sb.AppendLine(@"          <Paragraph>");
            sb.AppendLine(@"            <TextRuns>");
            sb.AppendLine(@"              <TextRun>");
            sb.AppendLine(@"                <Value>=Now()</Value>");
            sb.AppendLine(@"                <Style>");
            sb.AppendLine(@"                  <FontSize>10pt</FontSize>");
            sb.AppendLine(@"                  <Direction>RTL</Direction>");
            sb.AppendLine(@"                  <TextAlign>Right</TextAlign>");
            sb.AppendLine(@"                </Style>");
            sb.AppendLine(@"              </TextRun>");
            sb.AppendLine(@"            </TextRuns>");
            sb.AppendLine(@"          </Paragraph>");
            sb.AppendLine(@"        </Paragraphs>");
            sb.AppendLine(@"      </Textbox>");

            // ===== Table (Tablix) =====
            sb.AppendLine(@"      <Tablix Name=""Tablix1"">");
            sb.AppendLine(@"        <Top>3cm</Top>");
            sb.AppendLine(@"        <TablixBody>");
            sb.AppendLine(@"          <TablixColumns>");

            // 🔹 Columns (reversed order for RTL)
            for (int i = dt.Columns.Count - 1; i >= 0; i--)
            {
                sb.AppendLine(@"            <TablixColumn>");
                sb.AppendLine(@"              <Width>3cm</Width>");
                sb.AppendLine(@"            </TablixColumn>");
            }
            sb.AppendLine(@"          </TablixColumns>");

            sb.AppendLine(@"          <TablixRows>");
            // Header row
            sb.AppendLine(@"            <TablixRow>");
            sb.AppendLine(@"              <Height>0.8cm</Height>");
            sb.AppendLine(@"              <TablixCells>");

            for (int i = dt.Columns.Count - 1; i >= 0; i--)
            {
                string colName = dt.Columns[i].ColumnName;
                string headerText = GetHeaderText(columnMap, colName);
                sb.AppendLine(@"                <TablixCell>");
                sb.AppendLine(@"                  <CellContents>");
                sb.AppendLine(@"                    <Textbox Name=""Header_" + colName + @""">");
                sb.AppendLine(@"                      <Paragraphs>");
                sb.AppendLine(@"                        <Paragraph>");
                sb.AppendLine(@"                          <TextRuns>");
                sb.AppendLine(@"                            <TextRun>");
                sb.AppendLine(@"                              <Value>" + headerText + @"</Value>");
                sb.AppendLine(@"                              <Style>");
                sb.AppendLine(@"                                <FontSize>12pt</FontSize>");
                sb.AppendLine(@"                                <FontWeight>Bold</FontWeight>");
                sb.AppendLine(@"                                <Direction>RTL</Direction>");
                sb.AppendLine(@"                                <TextAlign>Right</TextAlign>");
                sb.AppendLine(@"                              </Style>");
                sb.AppendLine(@"                            </TextRun>");
                sb.AppendLine(@"                          </TextRuns>");
                sb.AppendLine(@"                        </Paragraph>");
                sb.AppendLine(@"                      </Paragraphs>");
                sb.AppendLine(@"                      <Style>");
                sb.AppendLine(@"                        <Border><Style>Solid</Style></Border>");
                sb.AppendLine(@"                      </Style>");
                sb.AppendLine(@"                    </Textbox>");
                sb.AppendLine(@"                  </CellContents>");
                sb.AppendLine(@"                </TablixCell>");
            }
            sb.AppendLine(@"              </TablixCells>");
            sb.AppendLine(@"            </TablixRow>");

            // Data row
            sb.AppendLine(@"            <TablixRow>");
            sb.AppendLine(@"              <Height>0.8cm</Height>");
            sb.AppendLine(@"              <TablixCells>");

            for (int i = dt.Columns.Count - 1; i >= 0; i--)
            {
                string colName = dt.Columns[i].ColumnName;
                sb.AppendLine(@"                <TablixCell>");
                sb.AppendLine(@"                  <CellContents>");
                sb.AppendLine(@"                    <Textbox Name=""Data_" + colName + @""">");
                sb.AppendLine(@"                      <Paragraphs>");
                sb.AppendLine(@"                        <Paragraph>");
                sb.AppendLine(@"                          <TextRuns>");
                sb.AppendLine(@"                            <TextRun>");
                sb.AppendLine(@"                              <Value>=Fields!" + colName + @".Value</Value>");
                sb.AppendLine(@"                              <Style>");
                sb.AppendLine(@"                                <FontSize>11pt</FontSize>");
                sb.AppendLine(@"                                <Direction>RTL</Direction>");
                sb.AppendLine(@"                                <TextAlign>Right</TextAlign>");
                sb.AppendLine(@"                              </Style>");
                sb.AppendLine(@"                            </TextRun>");
                sb.AppendLine(@"                          </TextRuns>");
                sb.AppendLine(@"                        </Paragraph>");
                sb.AppendLine(@"                      </Paragraphs>");
                sb.AppendLine(@"                      <Style>");
                sb.AppendLine(@"                        <Border><Style>Solid</Style></Border>");
                sb.AppendLine(@"                      </Style>");
                sb.AppendLine(@"                    </Textbox>");
                sb.AppendLine(@"                  </CellContents>");
                sb.AppendLine(@"                </TablixCell>");
            }
            sb.AppendLine(@"              </TablixCells>");
            sb.AppendLine(@"            </TablixRow>");
            sb.AppendLine(@"          </TablixRows>");
            sb.AppendLine(@"        </TablixBody>");

            // Column/Row Hierarchy
            sb.AppendLine(@"        <TablixColumnHierarchy><TablixMembers>");
            for (int i = dt.Columns.Count - 1; i >= 0; i--) sb.AppendLine(@"          <TablixMember />");
            sb.AppendLine(@"        </TablixMembers></TablixColumnHierarchy>");

            sb.AppendLine(@"        <TablixRowHierarchy><TablixMembers>");
            sb.AppendLine(@"          <TablixMember />"); // header
            sb.AppendLine(@"          <TablixMember><Group Name=""DetailGroup"" /></TablixMember>");
            sb.AppendLine(@"        </TablixMembers></TablixRowHierarchy>");

            sb.AppendLine(@"        <DataSetName>" + datasetName + @"</DataSetName>");
            sb.AppendLine(@"      </Tablix>");

            sb.AppendLine(@"    </ReportItems>");
            sb.AppendLine(@"    <Height>25cm</Height>");
            sb.AppendLine(@"  </Body>");

            // ---------------- Page Settings ----------------
            sb.AppendLine(@"  <Width>21cm</Width>");
            sb.AppendLine(@"  <Page>");
            sb.AppendLine(@"    <PageHeight>29.7cm</PageHeight>");
            sb.AppendLine(@"    <PageWidth>21cm</PageWidth>");
            sb.AppendLine(@"    <LeftMargin>1cm</LeftMargin>");
            sb.AppendLine(@"    <RightMargin>1cm</RightMargin>");
            sb.AppendLine(@"    <TopMargin>1cm</TopMargin>");
            sb.AppendLine(@"    <BottomMargin>1cm</BottomMargin>");
            sb.AppendLine(@"  </Page>");

            sb.AppendLine(@"</Report>");

            return new MemoryStream(Encoding.UTF8.GetBytes(sb.ToString()));
        }

        private static string GetHeaderText(Dictionary<string, string> columnMap, string colName)
        {
            foreach (var kv in columnMap)
            {
                if (kv.Value == colName) return kv.Key;
            }
            return colName; // fallback
        }
    }
}
