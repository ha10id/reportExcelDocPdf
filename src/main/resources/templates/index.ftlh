
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Organizer</title>
      <link rel="stylesheet" type="text/css" href="/css/style.css"/>
</head>
<body>
<h1>Organizer</h1>
<h3>Add task to list</h3>
<form action="/add" method="post">
    <table border="1px">
        <tr>
            <td>
                <label>Priority</label>
                <input type="number" name="priorityId">
            </td>
            <td>
                <label>Description</label>
                <input type="text" name="description">
            </td>
            <td>
                <label>Date</label>
                <input type="date" name="date">
            </td>

            <td>
                <input type="submit" value="add task">
            </td>

        </tr>
    </table>
</form>

<form action="/adds" method="post">
    <button type="submit">Add tasks</button>
</form>

<br>
<h2>Task list(task count : ${taskSize})</h2>
<table border="1px">
    <tr>
        <td>Id</td>
        <td>Priority</td>
        <td>Task description</td>
        <td>Date</td>
        <td>Delete from list</td>
    </tr>
    <#list taskList as task>
    <tr>
        <td>${task.id}</td>
        <td>${task.priorityId}</td>
        <td>${task.description}</td>
        <td>${task.date}</td>
        <td><a href="${'/delete/'+task.id}"">Delete</a></td>
    </tr>
    </#list>
</table>

<form action="/xls-to-pdf" method="post">
    <button type="submit">convert xlsx to doc & pdf</button>
</form>
<form action="/docx-to-pdf" method="post">
    <button type="submit">convert docx to pdf & pdf</button>
</form>
</body>
</html>
