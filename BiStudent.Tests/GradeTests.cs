using BiStudent.Contracts;
using BiStudent.Server.Services;
using ClosedXML.Excel;
using Xunit;

namespace BiStudent.Tests;

public sealed class GradeTests
{
    [Fact]
    public async Task CameraMarksPresentAfterLessonStart()
    {
        using var test = new TestDb();
        var teacher = test.AddTeacher();
        var group = test.AddGroup("ИВТ-1", "Иван");
        var started = DateTime.Now.AddMinutes(-30);
        var slot = test.AddSlot(teacher, group, start: TimeOnly.FromDateTime(started));
        await test.Attendance.MarkAsync(slot.Id, DateOnly.FromDateTime(started), group.Students[0].Id,
            AttendanceStatus.Absent, MarkMethod.Camera);
        Assert.Equal(AttendanceStatus.Present, Assert.Single(test.Db.AttendanceMarks).Status);
    }

    [Fact]
    public async Task ExcusedDefaultsReasonAndPresentClearsDetails()
    {
        using var test = new TestDb();
        var teacher = test.AddTeacher();
        var group = test.AddGroup("ИВТ-1", "Иван");
        var slot = test.AddSlot(teacher, group);
        var studentId = group.Students[0].Id;
        var date = new DateOnly(2026, 9, 1);
        await test.Attendance.MarkAsync(slot.Id, date, studentId, AttendanceStatus.Excused, MarkMethod.Manual,
            comment: "  Причина  ", hasDocument: true);
        var mark = Assert.Single(test.Db.AttendanceMarks);
        Assert.Equal(ExcuseReason.Other, mark.ExcuseReason);
        Assert.Equal("Причина", mark.Comment);
        await test.Attendance.MarkAsync(slot.Id, date, studentId, AttendanceStatus.Present, MarkMethod.Manual,
            ExcuseReason.Illness, "Не сохранять", true);
        Assert.Null(mark.ExcuseReason);
        Assert.Null(mark.Comment);
        Assert.False(mark.HasDocument);
    }

    [Fact]
    public async Task LessonGradeRejectsOutOfRangeAndNullDeletes()
    {
        using var test = new TestDb();
        var teacher = test.AddTeacher();
        var group = test.AddGroup("ИВТ-1", "Иван");
        var slot = test.AddSlot(teacher, group);
        var date = new DateOnly(2026, 9, 1);
        var studentId = group.Students[0].Id;
        await Assert.ThrowsAsync<InvalidOperationException>(() => test.Attendance.SetLessonGradeAsync(slot.Id, date, studentId, 0));
        await Assert.ThrowsAsync<InvalidOperationException>(() => test.Attendance.SetLessonGradeAsync(slot.Id, date, studentId, 101));
        await test.Attendance.SetLessonGradeAsync(slot.Id, date, studentId, 85);
        Assert.Equal(85, Assert.Single(test.Db.LessonGrades).Score);
        await test.Attendance.SetLessonGradeAsync(slot.Id, date, studentId, null);
        Assert.Empty(test.Db.LessonGrades);
    }

    [Fact]
    public async Task LessonGradeCreatesPresentMark()
    {
        using var test = new TestDb();
        var teacher = test.AddTeacher();
        var group = test.AddGroup("ИВТ-1", "Иван");
        var slot = test.AddSlot(teacher, group);
        var date = new DateOnly(2026, 9, 1);
        var studentId = group.Students[0].Id;
        var notifications = 0;
        test.Notifier.Changed += (slotId, changedDate) => notifications++;
        await test.Attendance.SetLessonGradeAsync(slot.Id, date, studentId, 85);
        var mark = Assert.Single(test.Db.AttendanceMarks);
        Assert.Equal(AttendanceStatus.Present, mark.Status);
        Assert.Equal(MarkMethod.Manual, mark.Method);
        Assert.Null(mark.ExcuseReason);
        Assert.Null(mark.Comment);
        Assert.False(mark.HasDocument);
        Assert.Equal(85, Assert.Single(test.Db.LessonGrades).Score);
        Assert.Equal(1, notifications);
    }

    [Theory]
    [InlineData(AttendanceStatus.Absent)]
    [InlineData(AttendanceStatus.Excused)]
    public async Task LessonGradeChangesAbsenceToPresent(AttendanceStatus status)
    {
        using var test = new TestDb();
        var teacher = test.AddTeacher();
        var group = test.AddGroup("ИВТ-1", "Иван");
        var slot = test.AddSlot(teacher, group);
        var date = new DateOnly(2026, 9, 1);
        var studentId = group.Students[0].Id;
        await test.Attendance.MarkAsync(slot.Id, date, studentId, status, MarkMethod.Manual,
            ExcuseReason.Illness, "Причина", true);
        await test.Attendance.SetLessonGradeAsync(slot.Id, date, studentId, 85);
        var mark = Assert.Single(test.Db.AttendanceMarks);
        Assert.Equal(AttendanceStatus.Present, mark.Status);
        Assert.Equal(MarkMethod.Manual, mark.Method);
        Assert.Null(mark.ExcuseReason);
        Assert.Null(mark.Comment);
        Assert.False(mark.HasDocument);
    }

    [Fact]
    public async Task LessonGradeKeepsCameraPresentMark()
    {
        using var test = new TestDb();
        var teacher = test.AddTeacher();
        var group = test.AddGroup("ИВТ-1", "Иван");
        var slot = test.AddSlot(teacher, group);
        var date = new DateOnly(2026, 9, 1);
        var studentId = group.Students[0].Id;
        await test.Attendance.MarkAsync(slot.Id, date, studentId,
            AttendanceStatus.Present, MarkMethod.Camera);
        var mark = Assert.Single(test.Db.AttendanceMarks);
        var markedAt = mark.MarkedAt;
        await test.Attendance.SetLessonGradeAsync(slot.Id, date, studentId, 85);
        Assert.Same(mark, Assert.Single(test.Db.AttendanceMarks));
        Assert.Equal(AttendanceStatus.Present, mark.Status);
        Assert.Equal(MarkMethod.Camera, mark.Method);
        Assert.Equal(markedAt, mark.MarkedAt);
    }

    [Fact]
    public async Task RemovingLessonGradeKeepsMark()
    {
        using var test = new TestDb();
        var teacher = test.AddTeacher();
        var group = test.AddGroup("ИВТ-1", "Иван");
        var slot = test.AddSlot(teacher, group);
        var date = new DateOnly(2026, 9, 1);
        var studentId = group.Students[0].Id;
        await test.Attendance.SetLessonGradeAsync(slot.Id, date, studentId, 85);
        await test.Attendance.MarkAsync(slot.Id, date, studentId, AttendanceStatus.Excused, MarkMethod.Manual,
            ExcuseReason.Illness, "Причина", true);
        var mark = Assert.Single(test.Db.AttendanceMarks);
        var markedAt = mark.MarkedAt;
        await test.Attendance.SetLessonGradeAsync(slot.Id, date, studentId, null);
        Assert.Empty(test.Db.LessonGrades);
        Assert.Same(mark, Assert.Single(test.Db.AttendanceMarks));
        Assert.Equal(AttendanceStatus.Excused, mark.Status);
        Assert.Equal(MarkMethod.Manual, mark.Method);
        Assert.Equal(ExcuseReason.Illness, mark.ExcuseReason);
        Assert.Equal("Причина", mark.Comment);
        Assert.True(mark.HasDocument);
        Assert.Equal(markedAt, mark.MarkedAt);
    }

    [Fact]
    public async Task RepeatedLessonGradeMarksAbsentStudentPresent()
    {
        using var test = new TestDb();
        var teacher = test.AddTeacher();
        var group = test.AddGroup("ИВТ-1", "Иван");
        var slot = test.AddSlot(teacher, group);
        var date = new DateOnly(2026, 9, 1);
        var studentId = group.Students[0].Id;
        await test.Attendance.SetLessonGradeAsync(slot.Id, date, studentId, 85);
        await test.Attendance.MarkAsync(slot.Id, date, studentId,
            AttendanceStatus.Absent, MarkMethod.Manual);
        var notifications = 0;
        test.Notifier.Changed += (slotId, changedDate) => notifications++;
        await test.Attendance.SetLessonGradeAsync(slot.Id, date, studentId, 85);
        Assert.Equal(85, Assert.Single(test.Db.LessonGrades).Score);
        Assert.Equal(AttendanceStatus.Present, Assert.Single(test.Db.AttendanceMarks).Status);
        Assert.Equal(MarkMethod.Manual, Assert.Single(test.Db.AttendanceMarks).Method);
        Assert.Equal(1, notifications);
    }

    [Fact]
    public async Task JournalCombinesLessonsAssessmentsAndFinal()
    {
        using var test = new TestDb();
        var teacher = test.AddTeacher();
        var group = test.AddGroup("ИВТ-1", "Иван");
        var slot = test.AddSlot(teacher, group);
        var studentId = group.Students[0].Id;
        var first = new DateOnly(2026, 9, 1);
        var second = first.AddDays(7);
        await test.Attendance.MarkAsync(slot.Id, first, studentId, AttendanceStatus.Present, MarkMethod.Manual);
        await test.Attendance.SetLessonGradeAsync(slot.Id, first, studentId, 85);
        await test.Attendance.MarkAsync(slot.Id, second, studentId, AttendanceStatus.Excused, MarkMethod.Manual,
            ExcuseReason.Illness);
        var assessment = await test.Grades.SaveAsync(teacher.Id, null,
            new SaveAssessmentRequest(group.Id, "Алгебра", "Тест", AssessmentKind.Test, first));
        await test.Grades.SetGradeAsync(teacher.Id, assessment.Id, studentId, 95);
        await test.Grades.SetFinalAsync(teacher.Id, group.Id, "Алгебра", studentId, 90);
        var journal = await test.Attendance.JournalAsync(teacher.Id, group.Id, first, second, "Алгебра");
        Assert.Equal(new[] { JournalColumnKind.Lesson, JournalColumnKind.Assessment, JournalColumnKind.Lesson },
            journal.Columns.Select(x => x.Kind));
        var row = Assert.Single(journal.Students);
        Assert.Equal(1, row.Present);
        Assert.Equal(0, row.Absent);
        Assert.Equal(1, row.Excused);
        Assert.Equal(2, row.MissedHours);
        Assert.Equal(2, row.ExcusedHours);
        Assert.Equal(50d, row.Percent);
        Assert.Equal(90d, row.Average);
        Assert.Equal(90, row.Final);
    }

    [Fact]
    public async Task AssessmentRejectsSubjectOutsideSchedule()
    {
        using var test = new TestDb();
        var teacher = test.AddTeacher();
        var group = test.AddGroup("ИВТ-1", "Иван");
        test.AddSlot(teacher, group);
        await Assert.ThrowsAsync<InvalidOperationException>(() => test.Grades.SaveAsync(teacher.Id, null,
            new SaveAssessmentRequest(group.Id, "Физика", "Тест", AssessmentKind.Test, new DateOnly(2026, 9, 1))));
    }

    [Fact]
    public async Task AssessmentWithGradesCannotChangeGroupOrSubject()
    {
        using var test = new TestDb();
        var teacher = test.AddTeacher();
        var group = test.AddGroup("ИВТ-1", "Иван");
        var other = test.AddGroup("ИВТ-2", "Анна");
        test.AddSlot(teacher, group);
        test.AddSlot(teacher, group, subject: "Физика");
        test.AddSlot(teacher, other);
        var date = new DateOnly(2026, 9, 1);
        var assessment = await test.Grades.SaveAsync(teacher.Id, null,
            new SaveAssessmentRequest(group.Id, "Алгебра", "Тест", AssessmentKind.Test, date));
        await test.Grades.SetGradeAsync(teacher.Id, assessment.Id, group.Students[0].Id, 85);
        var groupError = await Assert.ThrowsAsync<InvalidOperationException>(() => test.Grades.SaveAsync(teacher.Id, assessment.Id,
            new SaveAssessmentRequest(other.Id, "Алгебра", "Тест", AssessmentKind.Test, date)));
        Assert.Equal("Нельзя перенести мероприятие с выставленными баллами в другую группу.", groupError.Message);
        await Assert.ThrowsAsync<InvalidOperationException>(() => test.Grades.SaveAsync(teacher.Id, assessment.Id,
            new SaveAssessmentRequest(group.Id, "Физика", "Тест", AssessmentKind.Test, date)));
        Assert.Equal(group.Id, test.Db.Assessments.Single(x => x.Id == assessment.Id).GroupId);
    }

    [Fact]
    public async Task FillFinalRoundsAverageAndPreservesExisting()
    {
        using var test = new TestDb();
        var teacher = test.AddTeacher();
        var group = test.AddGroup("ИВТ-1", "Иван", "Анна");
        var slot = test.AddSlot(teacher, group);
        var date = new DateOnly(2026, 9, 1);
        var ivan = group.Students[0].Id;
        var anna = group.Students[1].Id;
        await test.Attendance.SetLessonGradeAsync(slot.Id, date, ivan, 84);
        await test.Attendance.SetLessonGradeAsync(slot.Id, date, anna, 60);
        var assessment = await test.Grades.SaveAsync(teacher.Id, null,
            new SaveAssessmentRequest(group.Id, "Алгебра", "Тест", AssessmentKind.Test, date));
        await test.Grades.SetGradeAsync(teacher.Id, assessment.Id, ivan, 85);
        await test.Grades.SetFinalAsync(teacher.Id, group.Id, "Алгебра", anna, 100);
        Assert.Equal(1, await test.Grades.FillFinalAsync(teacher.Id, group.Id, "Алгебра"));
        Assert.Equal(85, test.Db.FinalGrades.Single(x => x.StudentId == ivan).Score);
        Assert.Equal(100, test.Db.FinalGrades.Single(x => x.StudentId == anna).Score);
    }

    [Fact]
    public void ExportContainsScoresAndExcusedSymbol()
    {
        using var test = new TestDb();
        var matrix = new JournalMatrix(
            "Ирина Иванова", "ИВТ-1", "Алгебра", new DateOnly(2026, 9, 1), new DateOnly(2026, 9, 8), 2,
            [new JournalColumn(JournalColumnKind.Lesson, 1, new DateOnly(2026, 9, 1),
                new TimeOnly(9, 0), "Алгебра"),
             new JournalColumn(JournalColumnKind.Lesson, 1, new DateOnly(2026, 9, 8),
                new TimeOnly(9, 0), "Алгебра")],
            [new JournalStudent(1, "Иван Иванов",
                [new JournalCell(AttendanceStatus.Present, null, null, false, 85),
                 new JournalCell(AttendanceStatus.Excused, ExcuseReason.Illness, null, true, null)],
                1, 0, 1, 2, 2, 50, 85, null)]);
        using var stream = new MemoryStream(test.Export.JournalXlsx(matrix, "Журнал учёта посещаемости и успеваемости"));
        using var book = new XLWorkbook(stream);
        var sheet = book.Worksheet("Журнал");
        Assert.Equal("Журнал учёта посещаемости и успеваемости", sheet.Cell(1, 1).GetString());
        Assert.Equal("П 85", sheet.Cell(8, 3).GetString());
        Assert.Contains("Н(у)", test.Export.JournalCsv(matrix));
    }
}
