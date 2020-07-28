<!DOCTYPE html>
<html>

<head>
    <title>塔机轮压计算</title>
    <script type="text/javascript">
        function myCheck() {
            for (var i = 0; i < document.form1.elements.length - 1; i++) {
                if (document.form1.elements[i].value == "") {
                    alert("计算参数填写不完整");
                    document.form1.elements[i].focus();
                    return false;
                }
            }
            return true;

        }
    </script>

</head>

<body>
    <h1>行走塔机轮压计算</h1>
    <form name="form1" action="<?php echo htmlspecialchars($_SERVER["PHP_SELF"]); ?>" method="post" onsubmit="return myCheck()">
    <p>轮距(m):<input type="text" name="a" value="11"></p>
    <p>轨距(m):<input type="text" name="b" value="13.6"></p>
    <p>工作状态最大不平衡力矩(t.m):<input type="text" name="mw" value="1335"></p>
    <p>非工作状态最大不平衡力矩(t.m):<input type="text" name="mn" value="677"></p>
    <p>塔机自重(t):<input type="text" name="gt" value="365"></p>
    <p>行走机构自重(t):<input type="text" name="gx" value="90"></p>
    <p>最大吊重(t):<input type="text" name="g" value="70"></p>
    <p>备注:<input type="text" name="note" value="某某项目"></p>
    <p><input type="submit" value="计算"></p>
    </form>

    <?php
    if ($_SERVER["REQUEST_METHOD"] == "POST") {
        $a = $_POST["a"];
        $b = $_POST["b"];
        $mw = $_POST["mw"];
        $mn = $_POST["mn"];
        $gt = $_POST["gt"];
        $gx = $_POST["gx"];
        $g = $_POST["g"];
        $note = $_POST["note"];
        if ($a * $b * $mw * $mn * $gt * $gx * $g == 0) {
            echo "请检查输入数据<br>";
            #echo "<a href=\"index.php\">返回</a>";
            die();
        }

        
        #echo "<h2>计算参数</h2>";
        #echo "<p>轮距" . $a . "m<\p>";
        #echo "<p>轨距" . $b . "m<\p>";
        #echo "<p>工作状态最大不平衡力矩" . $mw . "t.m<\p>";
        #echo "<p>非工作状态最大不平衡力矩" . $mn . "t.m<\p>";
        #echo "<p>塔机自重" . $gt . "t<\p>";
        #echo "<p>行走机构自重" . $gx . "t<\p>";
        #echo "<p>最大吊重" . $g . "t<\p>";
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
        $max = max($fworkmax, $fnworkmax);
        $max = round($max, 1);
        $min = min($fnworkmin, $fnworkmin);
        $min = round($min, 1);
        echo "最大轮压为" . $max . "t<br>";
        echo "最小轮压为" . $min . "t<br>";
        echo "<br>";
        #echo "<a href=\"index.php\">返回</a>";

        $serverip = '127.0.0.1';
        $user = 'root';
        $passwd = 'xuming';
        $dbname = 'mycalc';
        $tbname = 'lunya';

        $conn = mysqli_connect($serverip, $user, $passwd);
        if (!$conn) {
            die('连接数据库失败' . mysqli_connect_error());
        }
        $sql = "CREATE DATABASE IF NOT EXISTS {$dbname} Character Set UTF8";
        mysqli_query($conn, $sql);
        mysqli_close($conn);
        $conn = mysqli_connect($serverip, $user, $passwd, $dbname);
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
    }
    ?>

<?php
    $serverip = '127.0.0.1';
    $user = 'root';
    $passwd = 'xuming';
    $dbname = 'mycalc';
    $tbname = 'lunya';
    $conn = mysqli_connect($serverip, $user, $passwd);
    if (!$conn) {
        die('无法连接数据库');
    }
    mysqli_query($conn, "set names utf8");
    $sql = "SELECT id,a,b,mw,mn,gt,gx,g,maxly,minly,note
        FROM {$tbname}";
    mysqli_select_db($conn, $dbname);
    $retval = mysqli_query($conn, $sql);
    if (!$retval) {
        die('无法读取数据: ' . mysqli_error($conn));
    }
    echo '<h2>计算记录</h2>';
    echo "\n";
    echo '<table border="1">
    <tr>
    <td>序号</td>
    <td>轮距(m)</td>
    <td>轨距(m)</td>
    <td>工作状态<br>最大不平衡力矩(t.m)</td>
    <td>非工作状态<br>最大不平衡力矩(t.m)</td>
    <td>塔机自重(t)</td>
    <td>行走机构自重(t)</td>
    <td>最大吊重(t)</td>
    <td>备注</td>
    <td>最大轮压(t)</td>
    <td>最小轮压(t)</td>
    </tr>';
    while ($row = mysqli_fetch_array($retval, MYSQLI_ASSOC)) {
        echo "<tr>
        <td>{$row['id']}</td><td>{$row['a']}</td><td>{$row['b']}</td><td>{$row['mw']}</td><td>{$row['mn']}</td><td>{$row['gt']}</td><td>{$row['gx']}</td><td>{$row['g']}</td><td>{$row['note']}</td><td>{$row['maxly']}</td><td>{$row['minly']}</td> 
        </tr>";
    }
    echo '</table>';
    mysqli_free_result($retval);
    mysqli_close($conn);
    echo "\n";

    ?>
</body>

</html>