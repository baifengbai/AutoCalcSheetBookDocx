<!DOCTYPE html>
<html>

<body>
    <?php
    
    $a = $_POST["a"];
    $b = $_POST["b"];
    $mw = $_POST["mw"];
    $mn = $_POST["mn"];
    $gt = $_POST["gt"];
    $gx = $_POST["gx"];
    $g = $_POST["g"];
    $note = $_POST["note"];

    if($a*$b*$mw*$mn*$gt*$gx*$g==0)
    {
        echo "请检查输入数据<br>";
        echo "<a href=\"index.php\">返回</a>";
        die();
    }

    echo "<h1>行走塔机轮压计算</h1>";
    echo "<h2>计算参数</h2>";
    echo "轮距" . $a . "m<br>";
    echo "轨距" . $b . "m<br>";
    echo "工作状态最大不平衡力矩" . $mw . "t.m<br>";
    echo "非工作状态最大不平衡力矩" . $mn . "t.m<br>";
    echo "塔机自重" . $gt . "t<br>";
    echo "行走机构自重" . $gx . "t<br>";
    echo "最大吊重" . $g . "t<br>";
    $l = $a * $b / (($a * $a + $b * $b) ** 0.5);
    $fworkmax = ($gt + $gx + $g) / 4 + $mw / (2 * $l);
    $fworkmin = ($gt + $gx + $g) / 4 - $mw / (2 * $l);
    echo "<h2>工作状态轮压计算</h2>";
    echo "工作状态最大轮压为" . round($fworkmax, 1) . "t<br>";
    echo "工作状态最小轮压为" . round($fworkmin, 1) . "t<br>";
    $mst = ($gt + $gx + $g) * min($a, $b) / 2;
    echo "工作状态稳定力矩" . round($mst, 1) . "t.m，不平衡力矩" . $mw . "t.m，稳定系数" . round($mst / $mw, 1) . "<br>";
    $fnworkmax = ($gt + $gx) / 4 + $mn / (2 * $l);
    $fnworkmin = ($gt + $gx) / 4 - $mn / (2 * $l);
    echo "<h2>非工作状态轮压计算</h2>";
    echo "非工作状态最大轮压为" . round($fnworkmax, 1) . "t<br>";
    echo "非工作状态最小轮压为" . round($fnworkmin, 1) . "t<br>";
    $mnst = ($gt + $gx) * min($a, $b) / 2;
    echo "非工作状态稳定力矩" . round($mnst, 1) . "t.m，不平衡力矩" . $mn . "t.m，稳定系数" . round($mnst / $mn, 1) . "<br>";
    
    echo "<h2>汇总结果</h2>";
    $max = max($fworkmax,$fnworkmax);
    $max = round($max, 1);
    $min = min($fnworkmin,$fnworkmin);
    $min = round($min, 1);
    echo "最大轮压为" . $max . "t<br>";
    echo "最小轮压为" . $min . "t<br>";
    echo "<br>";
    echo "<a href=\"index.php\">返回</a>";

    $serverip = '127.0.0.1';
    $user = 'root';
    $passwd = 'xuming';
    $dbname = 'mycalc';
    $tbname = 'lunya';

    $conn = mysqli_connect($serverip,$user,$passwd);
    if(!$conn){die('连接数据库失败'.mysqli_connect_error());}
    $sql = "CREATE DATABASE IF NOT EXISTS {$dbname} Character Set UTF8";
    mysqli_query($conn, $sql);
    mysqli_close($conn);
    $conn = mysqli_connect($serverip,$user,$passwd,$dbname);
    $sql = "CREATE TABLE IF NOT EXISTS {$tbname} (
        id INT(6) UNSIGNED AUTO_INCREMENT PRIMARY KEY, 
        a FLOAT(10) NOT NULL,
        b FLOAT(10) NOT NULL,
        mw FLOAT(10) NOT NULL,
        mn FLOAT(10) NOT NULL,
        gt FLOAT(10) NOT NULL,
        gx FLOAT(10) NOT NULL,
        g FLOAT(10) NOT NULL,
        maxly FLOAT(10) NOT NULL,
        minly FLOAT(10) NOT NULL,
        note TEXT NOT NULL
        )";
    mysqli_query($conn, $sql);
    $sql = "INSERT INTO {$tbname} (a,b,mw,mn,gt,gx,g,maxly,minly,note) 
            VALUES ($a,$b,$mw,$mn,$gt,$gx,$g,$max,$min,'$note')";
    mysqli_query($conn, $sql);
    mysqli_close($conn);

    ?>
</body>

</html>