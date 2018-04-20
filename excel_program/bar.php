

<!-- Styles -->
<style>
#<?php echo "work$worksheetNbr";?> {
	width	: 100%;
	height	: 500px;
}								
</style>

<!-- Resources -->
<script src="https://www.amcharts.com/lib/3/amcharts.js"></script>
<script src="https://www.amcharts.com/lib/3/serial.js"></script>
<script src="https://www.amcharts.com/lib/3/plugins/export/export.min.js"></script>
<link rel="stylesheet" href="https://www.amcharts.com/lib/3/plugins/export/export.css" type="text/css" media="all" />
<script src="https://www.amcharts.com/lib/3/themes/none.js"></script>

<!-- Chart code -->
<script>
var chart = AmCharts.makeChart("<?php echo "work$worksheetNbr";?>",
{
    "type": "serial",
    "theme": "light",
    "dataProvider": [
		<?php
foreach($AccountFieldOnChartBar as $o => $w){
echo '
{
        "name":"'.$o.'",
        "points":"'.$w.'"
		  },
  ';
 } 
 ?>
	
	],
    "valueAxes": [{
        "maximum": 168,
        "minimum": 0,
        "axisAlpha": 0,
        "dashLength": 4,
        "position": "left"
    }],
    "startDuration": 1,
    "graphs": [{
        "balloonText": "<span style='font-size:13px;'>[[category]]: <b>[[value]]</b></span>",
        "bulletOffset": 10,
        "bulletSize": 52,
        "colorField": "color",
        "cornerRadiusTop": 8,
        "customBulletField": "bullet",
        "fillAlphas": 0.8,
        "lineAlpha": 0,
        "type": "column",
        "valueField": "points"
    }],
    "marginTop": 0,
    "marginRight": 0,
    "marginLeft": 0,
    "marginBottom": 0,
    "autoMargins": false,
    "categoryField": "name",
    "categoryAxis": {
        "axisAlpha": 0,
        "gridAlpha": 0,
        "inside": true,
        "tickLength": 0
    },
    "export": {
    	"enabled": true
     }
});
</script>

<!-- HTML -->
<div id="<?php echo "work$worksheetNbr";?>"></div>
