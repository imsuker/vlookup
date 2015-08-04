var xlsx = require('node-xlsx');
~function(){
  var Dota = {
    config : {
      chartBaseExcel : {
        nameExcel : "base.xlsx",
        baseLine : 1,  //基于baseline行进行vlookup
        nMaxLine : 5,//从第5行开始写入
      },
      aTargetExcel : [
        {
          nameExcel : "a.xlsx",
          baseLine : 2, //基于baseline比对
          watchLine : 5  //观察值
        },
        {
          nameExcel : "b.xlsx",
          baseLine : 3, //基于baseline比对
          watchLine : 4  //观察值
        }

      ]
    },
    init : function(){
      Dota.run();
    },
    run : function(){
      var aChartBaseRow = xlsx.parse(Dota.config.chartBaseExcel.nameExcel)[0].data;
      var aTargetData = [];
      Dota.config.aTargetExcel.forEach(function(oTargetExcel, i){
        var oData = xlsx.parse(oTargetExcel.nameExcel)[0].data;
        aTargetData.push(oData);
      });

      var nMaxLine = Dota.config.chartBaseExcel.nMaxLine;
      aChartBaseRow.forEach(function(oRow, row){ //遍历基本表的每行
        if(row == 0){
          Dota.config.aTargetExcel.forEach(function(oTarget, i){
            oRow[nMaxLine-1 + i] = oTarget.nameExcel;
          });
          return;
        }
        var baseCell = oRow[Dota.config.chartBaseExcel.baseLine-1]; 
        Dota.config.aTargetExcel.forEach(function(oTargetExcel, i){//遍历目标表
          console.log('row', row,'i',i);
          var nTargetLine = nMaxLine -1 + i;
          var oTargetRow = aTargetData[i][row] || [];
          oTargetRow.some(function(aTargetRow, targetRow){//遍历目标表的每行
            var targetCell = aTargetRow[oTargetExcel.baseLine-1];
            console.log('target', oTargetExcel.nameExcel, 'row', row,'line', oTargetExcel.baseLine-1,'targetCell', targetCell);
            if(baseCell == targetCell){
              oRow[nTargetLine] = aTargetRow[oTargetExcel.watchLine-1];
              return true;
            } 
          });
        });
      });
      var buffer = xlsx.build([{name:"result", data : aChartBaseRow}]);
      var fs = require('fs');
      fs.writeFileSync('result.xlsx', buffer);
    },
    fun : null
  };
  Dota.run();
}();
