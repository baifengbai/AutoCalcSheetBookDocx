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
    <form name="form1" action="lunya.php" method="post" onsubmit="return myCheck()">
        轮距(m):<input type="text" name="a" value="11"><br>
        轨距(m):<input type="text" name="b" value="13.6"><br>
        工作状态最大不平衡力矩(t.m):<input type="text" name="mw" value="1335"><br>
        非工作状态最大不平衡力矩(t.m):<input type="text" name="mn" value="677"><br>
        塔机自重(t):<input type="text" name="gt" value="365"><br>
        行走机构自重(t):<input type="text" name="gx" value="90"><br>
        最大吊重(t):<input type="text" name="g" value="70"><br>
        备注:<input type="text" name="note" value="某某项目"><br>
        <input type="submit" value="计算">
    </form>
    <?php
    $serverip = '127.0.0.1';
    $user = 'root';
    $passwd = 'xuming';
    $dbname = 'mycalc';
    $tbname = 'lunya';
    $conn = mysqli_connect($serverip,$user,$passwd);
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