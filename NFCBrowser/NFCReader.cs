using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using PCSC;
using PCSC.Iso7816;

namespace NFCBrowser
{
    class NFCReader
    {
        DataTable dt = new DataTable();
        String readerName = "";
        
        public NFCReader()
        {
            dt.Columns.Add("Nama", Type.GetType("System.String"));
            dt.Columns.Add("Tanmatu", Type.GetType("System.String"));
            dt.Columns.Add("Proc", Type.GetType("System.String"));
            dt.Columns.Add("Date", Type.GetType("System.String"));
            dt.Columns.Add("IriSen", Type.GetType("System.String"));
            dt.Columns.Add("Iri", Type.GetType("System.String"));
            dt.Columns.Add("DeSen", Type.GetType("System.String"));
            dt.Columns.Add("De", Type.GetType("System.String"));
            dt.Columns.Add("Zandaka", Type.GetType("System.String"));
            dt.Columns.Add("Renban", Type.GetType("System.String"));
            dt.Columns.Add("Shiharai", Type.GetType("System.String"));
        }

        //
        // 正常な場合は、NFCReaderの名称を返す。なかったら、例外throw
        // 
        public String GetFirstReaderName()
        {
            using (var ctx = ContextFactory.Instance.Establish(SCardScope.User))
            {

                var firstReader = ctx.GetReaders().FirstOrDefault();

                if (firstReader == null)
                {
                    throw new Exception("カードリーダーが接続されてない");
                }

                using (var reader = ctx.ConnectReader(firstReader, SCardShareMode.Direct, SCardProtocol.Unset))
                {
                    var status = reader.GetStatus();
                    readerName = status.GetReaderNames()[0];
                }
            }
            return readerName;
        }
        // 
        // 正常な場合は、UIDを返す。なかったら、例外throw
        // 
        public String GetUID()
        {
            String UID = "";
            var contextFactory = ContextFactory.Instance;

            using (var context = contextFactory.Establish(SCardScope.System))
            {
                if ( this.readerName == "")
                {
                    this.readerName = GetFirstReaderName();
                }
                // 'using' statement to make sure the reader will be disposed (disconnected) on exit
                using (var rfidReader = context.ConnectReader(this.readerName, SCardShareMode.Shared, SCardProtocol.Any))
                {
                    var apdu = new CommandApdu(IsoCase.Case2Short, rfidReader.Protocol)
                    {
                        CLA = 0xFF,
                        Instruction = InstructionCode.GetData,
                        P1 = 0x00,
                        P2 = 0x00,
                        Le = 0 // We don't know the ID tag size
                    };

                    using (rfidReader.Transaction(SCardReaderDisposition.Leave))
                    {
                        var sendPci = SCardPCI.GetPci(rfidReader.Protocol);
                        var receivePci = new SCardPCI(); // IO returned protocol control information.

                        var receiveBuffer = new byte[256];
                        var command = apdu.ToArray();

                        var bytesReceived = rfidReader.Transmit(
                            sendPci, // Protocol Control Information (T0, T1 or Raw)
                            command, // command APDU
                            command.Length,
                            receivePci, // returning Protocol Control Information
                            receiveBuffer,
                            receiveBuffer.Length); // data buffer

                        var responseApdu = new ResponseApdu(receiveBuffer, bytesReceived, IsoCase.Case2Short, rfidReader.Protocol);
                        if (responseApdu.HasData)
                        {
                            UID = BitConverter.ToString(responseApdu.GetData());
                        }
                        else
                        {
                            throw new Exception ("UIDが取得できない");
                        }
                    }
                    
                }
                return UID;
            }
            
        }
        // 
        // 正常な場合は、履歴データのDataTableを返す。なかったら、例外throw
        // 
        public DataTable GetCardHistory()
        {
            // 履歴読み込み
            //
            
            var contextFactory = ContextFactory.Instance;

            using (var context = contextFactory.Establish(SCardScope.System))
            {
                if (this.readerName == "")
                {
                    this.readerName = GetFirstReaderName();
                }
                // 'using' statement to make sure the reader will be disposed (disconnected) on exit
                using (var rfidReader = context.ConnectReader(readerName, SCardShareMode.Shared, SCardProtocol.Any))
                {
                    byte[] dataIn = { 0x0f, 0x09 };

                    var apduSelectFile = new CommandApdu(IsoCase.Case4Short, rfidReader.Protocol)
                    {
                        CLA = 0xFF,
                        Instruction = InstructionCode.SelectFile,
                        P1 = 0x00,
                        P2 = 0x01,
                        // Lcは自動計算
                        Data = dataIn,
                        Le = 0 // 
                    };

                    using (rfidReader.Transaction(SCardReaderDisposition.Leave))
                    {
                        var sendPci = SCardPCI.GetPci(rfidReader.Protocol);
                        var receivePci = new SCardPCI(); // IO returned protocol control information.

                        var receiveBuffer = new byte[256];
                        var command = apduSelectFile.ToArray();

                        var bytesReceivedSelectedFile = rfidReader.Transmit(
                            sendPci, // Protocol Control Information (T0, T1 or Raw)
                            command, // command APDU
                            command.Length,
                            receivePci, // returning Protocol Control Information
                            receiveBuffer,
                            receiveBuffer.Length); // data buffer

                        var responseApdu = new ResponseApdu(receiveBuffer, bytesReceivedSelectedFile, IsoCase.Case2Short, rfidReader.Protocol);
                        var stationCodeDB = new SQLiteDB();
                        stationCodeDB.ConnectDB();
                        for (int i = 0; i < 20; ++i)
                        {

                            //② ReadBinaryとブロック指定
                            //176 = 0xB0
                            var apduReadBinary = new CommandApdu(IsoCase.Case2Short, rfidReader.Protocol)
                            {
                                CLA = 0xFF,
                                Instruction = InstructionCode.ReadBinary,
                                P1 = 0x00,
                                P2 = (byte)i,
                                Le = 0 // 
                            };

                            var commandReadBinary = apduReadBinary.ToArray();

                            var bytesReceivedReadBinary2 = rfidReader.Transmit(
                                sendPci, // Protocol Control Information (T0, T1 or Raw)
                                commandReadBinary, // command APDU
                                commandReadBinary.Length,
                                receivePci, // returning Protocol Control Information
                                receiveBuffer,
                                receiveBuffer.Length); // data buffer

                            var responseApdu2 =
                                new ResponseApdu(receiveBuffer, bytesReceivedReadBinary2, IsoCase.Case2Extended, rfidReader.Protocol);

                            // テストコード(データ解析関数の代わりに適当なDataTable作成)
                            // DataRow row = dt.NewRow();
                            // row["Nama"] = BitConverter.ToString(receiveBuffer, 0, 18);
                            // dt.Rows.Add(row);

                            // データ解析関数を実行
                            parse_tagWIthDB(receiveBuffer,stationCodeDB);

                            //parse_tag(receiveBuffer);
                        }
                        stationCodeDB.CloseDB();
                    }

                }
            }
            
            return this.dt;
        }
        //
        // データ解析関数 DataTableにデータを収める
        // 
        /*
        private void parse_tag(byte[] data)
        {
            // 年月日　+4〜+5 (2バイト): 年月日 [年/7ビット、月/4ビット、日/5ビット]
            // 入場地域コード
            // 入場線区コード
            // 入場駅コード
            // 出場地域コード
            // 出場線区コード
            // 出場駅コード
            // 残額
            
            int ctype, proc, date, time, balance, seq, region;
            int in_line, in_sta, out_line, out_sta;
            int yy, mm, dd;
            DataRow row = dt.NewRow();

            ctype = data[0];            // 端末種
            proc = data[1];             // 処理
            date = (data[4] << 8) + data[5];        // 日付
            balance = (data[10] << 8) + data[11];   // 残高
            balance = ((balance) >> 8) & 0xff | ((balance) << 8) & 0xff00;
            seq = (data[12] << 24) + (data[13] << 16) + (data[14] << 8) + data[15];
            region = seq & 0xff;        // Region
            seq >>= 8;                  // 連番

            out_line = -1;
            out_sta = -1;
            time = -1;

            switch (ctype)
            {
                case 0xC7:  // 物販
                case 0xC8:  // 自販機          
                    time = (data[6] << 8) + data[7];
                    in_line = data[8];
                    in_sta = data[9];
                    break;

                case 0x05:  // 車載機
                    in_line = (data[6] << 8) + data[7];
                    in_sta = (data[8] << 8) + data[9];
                    break;

                default:
                    in_line = data[6];
                    in_sta = data[7];
                    out_line = data[8];
                    out_sta = data[9];
                    break;
            }

            row["Tanmatu"] = consoleType(ctype);
            row["Proc"] = procType(proc);
            // 日付
            yy = date >> 9;
            mm = (date >> 5) & 0xf;
            dd = date & 0x1f;
            row["Date"] = yy.ToString() + "/" + mm.ToString() + "/" + dd.ToString();

            // 時刻
            if (time > 0)
            {
                int hh = time >> 11;
                int min = (time >> 5) & 0x3f;
                // 時間データは不要
                // row["Date"] += hh.ToString() + ":" + min.ToString();
            }

            // 入り地域コードが、左2ビット、出地域コードが、右２ビット
            int in_region, out_region;
            in_region = region >> 6;
            out_region = region >> 4;
            out_region = out_region & 0x3;

            if (out_line != -1)
            {
                if (row["Proc"].ToString() == "運賃支払")
                {
                    // 駅名検索する。
                    var db = new SQLiteDB();
                    String[] dump;
                    dump = db.selectStation(in_region.ToString("D"), in_line.ToString("D"), in_sta.ToString("D"));
                    row["IriSen"] = dump[0];
                    row["Iri"] = dump[1];
                    dump = db.selectStation(out_region.ToString("D"), out_line.ToString("D"), out_sta.ToString("D"));
                    row["DeSen"] = dump[0];
                    row["de"] = dump[1];

                }
            }
            row["Zandaka"] = balance.ToString();
            row["Renban"] = seq.ToString();
            // 生データ
            row["Nama"] = BitConverter.ToString(data, 0, 18);

            dt.Rows.Add(row);
        }
        */
        //
        // データ解析関数 DataTableにデータを収める
        // 
        private void parse_tagWIthDB(byte[] data, SQLiteDB db)
        {
            // 年月日　+4〜+5 (2バイト): 年月日 [年/7ビット、月/4ビット、日/5ビット]
            // 入場地域コード
            // 入場線区コード
            // 入場駅コード
            // 出場地域コード
            // 出場線区コード
            // 出場駅コード
            // 残額

            int ctype, proc, date, time, balance, seq, region;
            int in_line, in_sta, out_line, out_sta;
            int yy, mm, dd;
            DataRow row = dt.NewRow();

            ctype = data[0];            // 端末種
            proc = data[1];             // 処理
            date = (data[4] << 8) + data[5];        // 日付
            balance = (data[10] << 8) + data[11];   // 残高
            balance = ((balance) >> 8) & 0xff | ((balance) << 8) & 0xff00;
            seq = (data[12] << 24) + (data[13] << 16) + (data[14] << 8) + data[15];
            region = seq & 0xff;        // Region
            seq >>= 8;                  // 連番

            out_line = -1;
            out_sta = -1;
            time = -1;

            switch (ctype)
            {
                case 0xC7:  // 物販
                case 0xC8:  // 自販機          
                    time = (data[6] << 8) + data[7];
                    in_line = data[8];
                    in_sta = data[9];
                    break;

                case 0x05:  // 車載端末
                    in_line = (data[6] << 8) + data[7];
                    in_sta = (data[8] << 8) + data[9];
                    out_line = 0;
                    break;

                default:
                    in_line = data[6];
                    in_sta = data[7];
                    out_line = data[8];
                    out_sta = data[9];
                    break;
            }

            row["Tanmatu"] = consoleType(ctype);
            row["Proc"] = procType(proc);
            // 日付
            yy = date >> 9;
            mm = (date >> 5) & 0xf;
            dd = date & 0x1f;
            row["Date"] = yy.ToString() + "/" + mm.ToString() + "/" + dd.ToString();

            // 時刻
            if (time > 0)
            {
                int hh = time >> 11;
                int min = (time >> 5) & 0x3f;
                // 時間データは不要
                // row["Date"] += hh.ToString() + ":" + min.ToString();
            }

            // 入り地域コードが、左2ビット、出地域コードが、右２ビット
            int in_region, out_region;
            in_region = region >> 6;
            out_region = region >> 4;
            out_region = out_region & 0x3;

            if (out_line != -1)
            {
                // 電車
                if (row["Proc"].ToString() == "運賃支払" )
                {
                    // 駅名検索する。
                    String[] dump;
                    dump = db.selectStationNameToConnectedDB(in_region.ToString("D"), in_line.ToString("D"), in_sta.ToString("D"));

                    row["IriSen"] = dump[0];
                    row["Iri"] = dump[1];
                    dump = db.selectStationNameToConnectedDB(out_region.ToString("D"), out_line.ToString("D"), out_sta.ToString("D"));
                    row["DeSen"] = dump[0];
                    row["de"] = dump[1];
                }
                // バス
                if (row["Proc"].ToString() == "バス")
                {
                    // 駅名検索する。
                    String[] dump;
                    dump = db.selectBusNameToConnectedDB( in_line.ToString("D"), in_sta.ToString("D"));

                    row["IriSen"] = dump[0];
                    row["Iri"] = dump[1];
                    row["DeSen"] = "";
                    row["de"] = "";

                }


            }
            row["Zandaka"] = balance.ToString();
            row["Renban"] = seq.ToString();
            // 生データ
            row["Nama"] = BitConverter.ToString(data, 0, 18);

            dt.Rows.Add(row);
        }
        private string consoleType(int ctype)
        {
            switch (ctype)
            {
                case 0x03: return "清算機";
                case 0x05: return "車載端末";
                case 0x08: return "券売機";
                case 0x12: return "券売機";
                case 0x16: return "改札機";
                case 0x17: return "簡易改札機";
                case 0x18: return "窓口端末";
                case 0x1a: return "改札端末";
                case 0x1b: return "携帯電話";
                case 0x1c: return "乗継清算機";
                case 0x1d: return "連絡改札機";
                case 0xc7: return "物販";
                case 0xc8: return "自販機";
            }
            return "???";
        }

        private string procType(int proc)
        {
            switch (proc)
            {
                case 0x01: return "運賃支払";
                case 0x02: return "チャージ";
                case 0x03: return "券購";
                case 0x04: return "清算";
                case 0x07: return "新規";
                case 0x0d: return "バス";
                case 0x0f: return "バス";
                case 0x14: return "オートチャージ";
                case 0x46: return "物販";
                case 0x49: return "入金";
                case 0xc6: return "物販(現金併用)";
            }
            return "???";
        }



    }
}
