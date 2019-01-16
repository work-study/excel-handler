<!DOCTYPE html>
<html lang="zh">
<head>
    <meta charset="UTF-8">
    <title>Document</title>
    <link href="/css/index.css" rel="stylesheet"/>
</head>
<body>
<div class="container">
    <form action="/excel" method="post" enctype="multipart/form-data">
        <table>
            <tr>
                <td>
                    <span>添加附件:</span>
                    <input id="fileUpload" type="file" name="file">
                </td>
            </tr>
            <tr>
                <td>
                    <span>那些列产生唯一值，逗号隔开:</span>
                    <input id="cellListStr" type="text" name="cellListStr">
                </td>
            </tr>
            <tr>
                <td>
                    <span>唯一值放到那一行:</span>
                    <input id="resultCell" type="text" name="resultCell">
                </td>
            </tr>
        </table>
        <input type="submit" value="执行">
    </form>

</div>
</body>
</html>