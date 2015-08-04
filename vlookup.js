var xlsx = require('node-xlsx');
~function(){
  var Dota = {
    config : {
      infoOfOriginExcel : { //源excel信息
        nameExcel : "base.xlsx",
        lineBase : 1,  //基于baseline行进行vlookup
        lineBeginInsert : 6,//从第5行开始写入
      },
      aInfoOfTargetExcels : [//目标excels信息
        {
          nameExcel : "a.xlsx",
          lineBase : 2, //基于lineBase比对
          lineWatch : 5  //观察值
        },
        {
          nameExcel : "b.xlsx",
          lineBase : 3, //基于baseline比对
          lineWatch : 4  //观察值
        }

      ]
    },
    init : function(){
      Dota.run();
    },
    run : function(){
      var aRowsBaseExcel = xlsx.parse(Dota.config.infoOfOriginExcel.nameExcel)[0].data; //baseExcel的数据
      //存放所有targetExcels的数据
      var aDataOfTargetExcels = [];
      Dota.config.aInfoOfTargetExcels.forEach(function(oInfoOfTargetExcel, i){
        var oData = xlsx.parse(oInfoOfTargetExcel.nameExcel)[0].data;
        aDataOfTargetExcels.push(oData);
      });

      //开始运算
      var lineBeginInsert = Dota.config.infoOfOriginExcel.lineBeginInsert;
      aRowsBaseExcel.forEach(function(oRowBaseExcel, rowOfBaseExcel){ //遍历基本表的每行
        if(rowOfBaseExcel == 0){//第一列是输出标题，不用运算
          Dota.config.aInfoOfTargetExcels.forEach(function(oTarget, i){
            oRowBaseExcel[lineBeginInsert-1 + i] = oTarget.nameExcel;
          });
          return;
        }
        //将要比较的值
        var valueOfCellBase = oRowBaseExcel[Dota.config.infoOfOriginExcel.lineBase-1]; 
        //遍历目标表找寻目标单元格
        Dota.config.aInfoOfTargetExcels.forEach(function(oInfoOfTargetExcel, i){//遍历目标表
          console.log('rowOfBaseExcel', rowOfBaseExcel,'targetExcel',oInfoOfTargetExcel.nameExcel,'valueOfCellBase', valueOfCellBase);
          var nLineWillInsert = lineBeginInsert -1 + i; //每行vlookup结果放置再该列
          var aAllDataOfTargetExcel = aDataOfTargetExcels[i] || [];//取出目标表的所有行
          var bMatch = false;
          aAllDataOfTargetExcel.some(function(aDataOfTargetRow, rowTarget){//遍历目标表的每行
            var valueOfCheckCellTargetLine = aDataOfTargetRow[oInfoOfTargetExcel.lineBase-1]; //目标表的目标列的值
            console.log('rowTarget', rowTarget,'lineTarget', oInfoOfTargetExcel.lineBase-1,'valueOfLineTarget', valueOfCheckCellTargetLine);
            if(valueOfCellBase == valueOfCheckCellTargetLine){//比较，如果相等，写入baseexcel的base行+base列
              var result = aDataOfTargetRow[oInfoOfTargetExcel.lineWatch-1];
              oRowBaseExcel[nLineWillInsert] = result;
              bMatch = true;
              console.log('match succsss!','nLineWillInsert', nLineWillInsert + 1,'lineWatch',oInfoOfTargetExcel.lineWatch-1 + 1, 'result', result);
            }
          });
          if(!bMatch){
              oRowBaseExcel[nLineWillInsert] = '#N/A';
          }
        });
      });
      var buffer = xlsx.build([{name:"result", data : aRowsBaseExcel}]);
      var fs = require('fs');
      fs.writeFileSync('result.xlsx', buffer);
    },
    fun : null
  };
  Dota.run();
}();
