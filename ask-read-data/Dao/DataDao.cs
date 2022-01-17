using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ask_read_data.Dao
{
    public class DataDao
    {
        public static string SP_BU_Mastar_GetAll2000Record()
        {
            var commandText = $@"  SELECT TOP (1000)  
	                               [WAYMD]
                                  ,[SEQ]
                                  ,[KATASIKI]
                                  ,[MEISHO]
                                  ,[FILLER1]
                                  ,[OPT]
                                  ,[JIKU]
                                  ,[FILLER2]
                                  ,[DAI]
                                  ,[MC]
                                  ,[SIMUKE]
                                  ,[E0]
                                  ,[BUBAN]
                                  ,[TANTO]
                                  ,[GR]
                                  ,[KIGO]
                                  ,[MAKR]
                                  ,[KOSUU]
                                  ,[KISYU]
                                  ,[MEWISYO]
                                  ,[FYMD]
                                  ,[SEIHINCD]
                                  ,[SEHINJNO]
                                  ,[FileName]
                                  ,[LineNumber]
                                  ,[Position]

	                               FROM [ask_datadb_test].[dbo].[DataImport]";

            return commandText;
        }

        public static string SP_BU_Mastar_SearchDataImport()
        {
            var commandText = $@"  SELECT TOP (2000)  
	                               [WAYMD]
                                  ,[SEQ]
                                  ,[KATASIKI]
                                  ,[MEISHO]
                                  ,[FILLER1]
                                  ,[OPT]
                                  ,[JIKU]
                                  ,[FILLER2]
                                  ,[DAI]
                                  ,[MC]
                                  ,[SIMUKE]
                                  ,[E0]
                                  ,[BUBAN]
                                  ,[TANTO]
                                  ,[GR]
                                  ,[KIGO]
                                  ,[MAKR]
                                  ,[KOSUU]
                                  ,[KISYU]
                                  ,[MEWISYO]
                                  ,[FYMD]
                                  ,[SEIHINCD]
                                  ,[SEHINJNO]
                                  ,[FileName]
                                  ,[LineNumber]
                                  ,[Position]

	                               FROM [ask_datadb_test].[dbo].[DataImport]
	                               WHERE FORMAT(CreateDateTime, 'yyyy-MM-dd 00:00:00') = FORMAT(@CreateDateTime, 'yyyy-MM-dd 00:00:00')";

            return commandText;
        }
    }
}
