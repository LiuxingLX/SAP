var databasePath = "C:/Users/liuxing/Desktop/test/src/database/data.mdb"

function initial_cnn() {
  var conn = new ActiveXObject("ADODB.Connection");
  connstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + databasePath+";Jet OLEDB:Database password=LCC;";
  conn.Open(connstr);
  return conn;
}

function updateData(dataArray){
  var conn = initial_cnn();
  var sql="update table1 set data1="+dataArray[0]+",data2="+dataArray[1]+",data3="+dataArray[2];
  try {
    conn.execute(sql);
    conn.close();
  } catch (e) {
    conn.close();
    console.log(e);
  }
}

function saveData(){
  var conn = initial_cnn();
  var dataArray = [oChart.series[0].data[0].y,oChart.series[0].data[1].y,oChart.series[0].data[2].y];
  var sql="insert into table2(data1,data2,data3) values("+dataArray[0]+","+dataArray[1]+","+dataArray[2]+")";
  try {
    conn.execute(sql);
    conn.close();
    fetchData();
  } catch (e) {
    conn.close();
    console.log(e);
  }
}
function sqlQuery(msql) {
  $('#tb2 tr').remove();
  var conn = initial_cnn();
  var rs = new ActiveXObject("ADODB.Recordset");
  if (msql == undefined) {
    var sql = $("#msql").html().replace("&gt;", ">").replace("&lt;", "<");
  } else {
    var sql = msql;
  }

  try {
    rs.open(sql, conn);
    var cCount = rs.Fields.count;
    var result = "<tr>";
    for (var i = 0; i < cCount; i++) {
      result += "<th>" + rs.Fields.item(i).name + "</th>";
    }
    result += "</tr>";
    $('#tb2').append(result);

    while (!rs.EOF) {
      var result = "";
      result = "<tr>";
      for (var i = 0; i < cCount; i++) {
        if (typeof(rs.Fields.item(i).value) == 'date') {
          result += "<td>" + getDateStr3(new Date(rs.Fields.item(i).value), '/') + "</td>";
        } else {
          result += "<td>" + rs.Fields.item(i).value + "</td>";
        }
      }
      result += "</tr>"
      $('#tb2').append(result);
      rs.moveNext();
    }
  } catch (e) {
    console.log(e);
    console.log("数据库出错");
    conn.close();
  }
  $('#table-count span').html($('#tb2 tr').length-1);
}

function fetchData(dataArray){
  sqlQuery("select top 100 * from table2");
}

//取随机数
function mRandom (m,n){
  var num = Math.floor(Math.random()*(m - n) + n);
  return num;
}

function changeData(){
  var tempdata=[mRandom(1,100),mRandom(1,100),mRandom(1,100)];
  container.SetData(tempdata);
  updateData(tempdata);
  fetchData();
}
function getsrc(){
  var temppath= location.href.split("///")[1];
  temppath=temppath.split("/");
  var finalpath="";
  for (var i = 0; i < temppath.length-1; i++) {
    finalpath+=temppath[i]+"/";
  }
  finalpath+="src/database/data.mdb"
  databasePath=finalpath;
  // alert(finalpath);
}
