using System.Net;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text.Json;
using System.Text.Json.Serialization;
using BiStudent.Contracts;
using BiStudent.Server.Data;
using BiStudent.Server.Services;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc.Testing;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.DependencyInjection;
using Xunit;

namespace BiStudent.Tests;

// Импорт расписания колледжа и списков групп из Excel.
public sealed class ImportTests : IDisposable
{
    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web) { Converters = { new JsonStringEnumConverter() } };
    private readonly string temporaryDirectory = Path.Combine(Path.GetTempPath(), $"BiStudent.Tests.{Guid.NewGuid():N}");
    private readonly WebApplicationFactory<Program> factory;
    private readonly HttpClient client;

    public ImportTests()
    {
        Directory.CreateDirectory(temporaryDirectory);
        factory = new ImportFactory(temporaryDirectory);
        client = factory.CreateClient();
    }

    /// Мини-копия формата расписания колледжа: две группы, объединённые ячейки дня и звонка, строки преподавателей.
    internal static byte[] CollegeSchedule()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("4 курс");
        sheet.Cell(3, 1).Value = "САБАҚ КЕСТЕСІ · 4 курс · 1 СЕМЕСТР (недели 1-18, 31.08.2026–01.01.2027)";
        foreach (var (column, group) in new[] { (5, "23-ПО-103-4р-1"), (10, "23-ПО-103-4р-2") })
        {
            sheet.Cell(6, column - 4).Value = "Күні Дни";
            sheet.Cell(6, column - 3).Value = "Қоңырау уақыты Время звонка";
            sheet.Cell(6, column - 2).Value = "No ; пары";
            sheet.Cell(6, column - 1).Value = "Дәрісхана ; Аудитория";
            sheet.Cell(6, column).Value = group;
            sheet.Cell(7, column - 4).Value = "Дүйсенбі / Понедельник";
            sheet.Range(7, column - 4, 10, column - 4).Merge();
            sheet.Cell(7, column - 3).Value = "08:00 09:20";
            sheet.Range(7, column - 3, 8, column - 3).Merge();
            sheet.Cell(9, column - 3).Value = "09:30 10:50";
            sheet.Range(9, column - 3, 10, column - 3).Merge();
            sheet.Cell(7, column - 2).Value = 1;
            sheet.Cell(9, column - 2).Value = 2;
        }
        sheet.Cell(7, 5).Value = "Искусственный интеллект OpenCV · 711 · нед. 6-7\n[ПРАКТИКА] База данных PL SQL/MONGODB · уточнить · нед. 3 (разово)";
        sheet.Cell(8, 5).Value = "Ниязова М.Е. · 711\nХамит А.Ү. · уточнить";
        sheet.Cell(9, 5).Value = "[СРОП] СРОП · Автоматизация DevOps · 504 · с 23-ПО-103-4р-2 · нед. 1 (разово), 5-7";
        sheet.Cell(10, 5).Value = "вакансия (Основы экономики и  предпринимательской деятельности) · 504";
        sheet.Cell(7, 10).Value = "[ПРАКТИКА] Программирование С# · 505 · нед. 2-4    Искусственный интеллект OpenCV -уточнить нед, 14-15";
        sheet.Cell(8, 10).Value = "Ниязова -505";
        sheet.Cell(9, 10).Value = "Непонятная запись без недель";
        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        return stream.ToArray();
    }

    [Fact]
    public void Parser_ReadsCollegeFormat()
    {
        using var stream = new MemoryStream(CollegeSchedule());
        var parsed = CollegeScheduleParser.Parse(stream, new DateOnly(2026, 9, 1));
        Assert.Equal(new DateOnly(2026, 8, 31), parsed.SemesterStart);
        Assert.Equal(["23-ПО-103-4р-1", "23-ПО-103-4р-2"], parsed.Groups.Select(x => x.Group));
        var first = parsed.Lessons.Where(x => x.Group == "23-ПО-103-4р-1").ToList();
        Assert.Contains(first, x => x is { Subject: "Искусственный интеллект OpenCV", Room: "711", Teacher: "Ниязова М.Е.", FromWeek: 6, ToWeek: 7, Once: false,
            Day: DayOfWeek.Monday } && x.Start == new TimeOnly(8, 0) && x.End == new TimeOnly(9, 20));
        Assert.Contains(first, x => x is { Subject: "База данных PL SQL/MONGODB", Room: null, Teacher: "Хамит А.Ү.", FromWeek: 3, ToWeek: 3, Once: true });
        // СРОП: дисциплина — второе поле, «с группой» — не аудитория; строка вакансии с двойным пробелом не рвётся.
        var srop = first.Where(x => x.Subject == "Автоматизация DevOps").ToList();
        Assert.Equal(2, srop.Count);
        Assert.All(srop, x => Assert.Equal(("504", CollegeScheduleParser.Vacancy, new TimeOnly(9, 30)), (x.Room, x.Teacher, x.Start)));
        Assert.Contains(srop, x => x is { FromWeek: 1, Once: true });
        Assert.Contains(srop, x => x is { FromWeek: 5, ToWeek: 7, Once: false });
        var second = parsed.Lessons.Where(x => x.Group == "23-ПО-103-4р-2").ToList();
        Assert.Contains(second, x => x is { Subject: "Программирование С#", Room: "505", FromWeek: 2, ToWeek: 4 });
        Assert.Contains(second, x => x is { Subject: "Искусственный интеллект OpenCV", Room: null, FromWeek: 14, ToWeek: 15 });
        // «Ниязова -505» сводится к единственной полной записи с той же фамилией.
        Assert.All(second, x => Assert.Equal("Ниязова М.Е.", x.Teacher));
        Assert.Single(parsed.Warnings);
        Assert.Contains("Непонятная запись", parsed.Warnings[0]);
    }

    [Theory]
    [InlineData("[ПРАКТИКА] Программирование С# · 505 · нед.5,6,7", "Программирование С#", "505", "5-7")]
    [InlineData("[ПРАКТИКА] Программирование С# -нед.5-7", "Программирование С#", null, "5-7")]
    [InlineData("[ПРАКТИКА] Программирование С# -5-7", "Программирование С#", null, "5-7")]
    [InlineData("Автоматизация DevOps · 711 · нед. 8 (разово)", "Автоматизация DevOps", "711", "8-8")]
    [InlineData("[ПРАКТИКА] Программирование С# ·504-нед. 9 (разово)", "Программирование С#", "504", "9-9")]
    [InlineData("[ПРАКТИКА] Программирование С# -уточн-нед. 8-9", "Программирование С#", null, "8-9")]
    [InlineData("[ЛЕКЦИЯ] Основы экономики · 804 · с 23-ПО-103-4р-1 · нед. 5-7, 14 (разово)", "Основы экономики", "804", "5-7,14-14")]
    [InlineData("Профессиональный иностранный язык-2 3 курс · 810 · нед.  9-11", "Профессиональный иностранный язык-2 3 курс", "810", "9-11")]
    [InlineData("[ПРАКТИКА] Аппаратное обеспечение компьютерной техники · уточнить-816 · нед.  5-7, 14-15", "Аппаратное обеспечение компьютерной техники", "816", "5-7,14-15")]
    [InlineData("Компьютерные сети · 812 · нед. 9 (разово),  18 (разово)", "Компьютерные сети", "812", "9-9,18-18")]
    [InlineData("Графика · 811 · нед.1,2,3,4,5,6,7,  14,15", "Графика", "811", "1-7,14-15")]
    [InlineData("Графика · 811 · нед.5,6,7,  14-16", "Графика", "811", "5-7,14-16")]
    [InlineData("[ФИЗКУЛЬТУРА] Физическая культура 3 курс · Спортзал (поток 1) · с 24-ЦТ-504-3р, 24-ПД-101-3р-2,  24-ПД-101-3р-3· нед. 5",
        "Физическая культура 3 курс", "Спортзал (поток 1)", "5-5")]
    [InlineData("Иностранный язык 2 курс · 808 · нед.5-7 -805", "Иностранный язык 2 курс", "805", "5-7")]
    [InlineData("[ПРАКТИКА] Программирование С# -5-7,", "Программирование С#", null, "5-7")]
    public void Parser_ToleratesHandEditedEntries(string entry, string subject, string? room, string weeks)
    {
        Assert.True(CollegeScheduleParser.TryParseEntry(entry, out var parsedSubject, out var parsedRoom, out var parsedWeeks));
        Assert.Equal(subject, parsedSubject);
        Assert.Equal(room, parsedRoom);
        Assert.Equal(weeks, string.Join(",", parsedWeeks.Select(x => $"{x.From}-{x.To}")));
    }

    [Theory]
    [InlineData("Профессиональный иностранный язык-2 3 курс · 810 · нед.  9-11")]
    [InlineData("[ПРАКТИКА] Аппаратное обеспечение компьютерной техники · уточнить-816 · нед.  5-7, 14-15")]
    [InlineData("Компьютерные сети · 812 · нед. 9 (разово),  18 (разово)")]
    [InlineData("Графика · 811 · нед.1,2,3,4,5,6,7,  14,15")]
    [InlineData("Графика · 811 · нед.5,6,7,  14-16")]
    [InlineData("[ФИЗКУЛЬТУРА] Физическая культура 3 курс · Спортзал (поток 1) · с 24-ЦТ-504-3р, 24-ПД-101-3р-2,  24-ПД-101-3р-3· нед. 5")]
    public void Parser_KeepsWeekContinuationsInOneEntry(string text) =>
        Assert.Single(CollegeScheduleParser.SplitEntries(text));

    [Fact]
    public void Parser_DoesNotGlueTaggedEntryToTrailingComma() =>
        Assert.Equal(["[ПРАКТИКА] Программирование С# · 505 · нед. 1 (разово), 3 (разово)", "[ПРАКТИКА] Программирование С# -5-7,",
                "[ПРАКТИКА] Программирование на языке высого уровня · 514 · нед. 18 (разово)"],
            CollegeScheduleParser.SplitEntries("[ПРАКТИКА] Программирование С# · 505 · нед. 1 (разово), 3 (разово)  [ПРАКТИКА] Программирование С# -5-7,     [ПРАКТИКА] Программирование на языке высого уровня · 514 · нед. 18 (разово)"));

    [Theory]
    [InlineData("1,2,3,4", "1-4:False")]
    [InlineData("5-7, 8", "5-8:False")]
    [InlineData("2-4, 5 (разово)", "2-4:False,5-5:True")]
    [InlineData("1 (разово), 2", "1-1:True,2-2:True")]
    [InlineData("8 (разово), 10-15, 18 (разово)", "8-8:True,10-15:False,18-18:True")]
    public void Parser_JoinsConsecutiveWeeksIntoRanges(string text, string expected) =>
        Assert.Equal(expected, string.Join(",", CollegeScheduleParser.ParseWeeks(text).Select(x => $"{x.From}-{x.To}:{x.Once}")));

    [Theory]
    [InlineData("[ЛЕКЦИЯ] Основы экономики и предпринимательской деятельности · 705")]
    [InlineData("математическая статистика 705, нед), 14-18, поток")]
    public void Parser_WarnsForEntriesWithoutWeeks(string text) =>
        Assert.False(CollegeScheduleParser.TryParseEntry(text, out _, out _, out _));

    [Fact]
    public void Parser_DropsBareSropWithoutShiftingTeachers()
    {
        using var source = new MemoryStream(CollegeSchedule());
        using var workbook = new XLWorkbook(source);
        var sheet = workbook.Worksheet("4 курс");
        sheet.Cell(7, 5).Value = "СРОП [ФИЗКУЛЬТУРА] Физическая культура 2 курс · Спортзал (поток 1) · нед. 5-7, 14-18\nГрафика · 811 · нед. 5\n[ЛЕКЦИЯ] Биология · 705 · нед. 5";
        sheet.Cell(8, 5).Value = "Битуреева Э.Т. · Спортзал (поток 1)\nАкбанов А.К. · 811\nАбдиманапова Г.С. · 705";
        sheet.Cell(9, 5).Value = "";
        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        stream.Position = 0;

        var parsed = CollegeScheduleParser.Parse(stream, new DateOnly(2026, 8, 31));
        var first = parsed.Lessons.Where(x => x.Group == "23-ПО-103-4р-1").ToList();
        Assert.DoesNotContain(parsed.Warnings, x => x.Contains("23-ПО-103-4р-1"));
        Assert.Contains(first, x => x is { Subject: "Физическая культура 2 курс", Room: "Спортзал (поток 1)", Teacher: "Битуреева Э.Т.", FromWeek: 5, ToWeek: 7 });
        Assert.Contains(first, x => x is { Subject: "Графика", Teacher: "Акбанов А.К." });
        Assert.Contains(first, x => x is { Subject: "Биология", Teacher: "Абдиманапова Г.С." });
    }

    [Theory]
    [InlineData("Асанова К.С. 806            Мансурова Г.М. · 514", "Асанова К.С.|Мансурова Г.М.")]
    [InlineData("Асанова К.С. -Қабылқай Ә.Қ. · Спортзал (поток 1)", "Асанова К.С.|Қабылқай Ә.Қ.")]
    [InlineData("Битуреева Э.Т. Думшебаева А.Т. · 804", "Битуреева Э.Т.|Думшебаева А.Т.")]
    [InlineData("Рахимжанова С.К. Уточнить    Канат Т. · 504", "Рахимжанова С.К.|Канат Т.")]
    [InlineData("Тайгара А--уточнить Ниязова М.Е. · 508\nНиязова М.Е. · 504", "Тайгара А.|Ниязова М.Е.|Ниязова М.Е.")]
    [InlineData("Хамит А.Ү. · 704\nХамит А.Ү. · 804\nХамит А.Ү. · 804   Ниязова Ц4, Ниязова -505",
        "Хамит А.Ү.|Хамит А.Ү.|Хамит А.Ү.|Ниязова|Ниязова")]
    [InlineData("Ниязова М.Е. · Ц2\nНиязова М.Е. · 812  Ниязова М.Е. 506\nНиязова М.Е. · 812",
        "Ниязова М.Е.|Ниязова М.Е.|Ниязова М.Е.|Ниязова М.Е.")]
    [InlineData("Ниязова М.Е. · 505,   Тайгара А-514", "Ниязова М.Е.|Тайгара А.")]
    [InlineData("Тайгара-уточнить", "Тайгара")]
    [InlineData("Тайгара А-уточнить", "Тайгара А.")]
    [InlineData("Әбікенова Ж.Ж.-уточнить", "Әбікенова Ж.Ж.")]
    [InlineData("Лесбек А.Л.· 705", "Лесбек А.Л.")]
    [InlineData("Рахимжанова С.К.- уточнить", "Рахимжанова С.К.")]
    [InlineData("Сариева А.К. -805", "Сариева А.К.")]
    [InlineData("вакансия (Основы экономики и  предпринимательской деятельности) · Ц4", "Вакансия")]
    [InlineData("Эрик (ментор) · Ц1", "Эрик (ментор)")]
    public void Parser_ExtractsTeacherNames(string text, string names) =>
        Assert.Equal(names.Split('|'), CollegeScheduleParser.ExtractTeachers(text));

    [Fact]
    public void Parser_UnifiesTeacherSpellingsAndIncompleteInitials()
    {
        var names = new[] { "Асанова", "Асанова К.С.", "Әбікенова Ж.", "Әбікенова Ж.Ж.",
            "Ниязова", "Ниязова М.Е.", "Рахимжанова С.К", "Рахимжанова С.К.",
            "Сайдинова Э.Б", "Сайдинова Э.Б.", "Султанов Д.", "Султанов Д.Б.",
            "Сұлтанов Д.Б.", "Сұлтанов Д.Б.", "Тайгара", "Тайгара А.", "Тайғара А.М.",
            "Тунгенбаева К.Б.", "Түнгенбаева К.Б.", "Түнгенбаева К.Б." };
        Assert.Equal(new[] { "Асанова К.С.", "Асанова К.С.", "Әбікенова Ж.Ж.", "Әбікенова Ж.Ж.",
            "Ниязова М.Е.", "Ниязова М.Е.", "Рахимжанова С.К.", "Рахимжанова С.К.",
            "Сайдинова Э.Б.", "Сайдинова Э.Б.", "Сұлтанов Д.Б.", "Сұлтанов Д.Б.",
            "Сұлтанов Д.Б.", "Сұлтанов Д.Б.", "Тайғара А.М.", "Тайғара А.М.", "Тайғара А.М.",
            "Түнгенбаева К.Б.", "Түнгенбаева К.Б.", "Түнгенбаева К.Б." },
            CollegeScheduleParser.UnifyTeacherNames(names));
    }

    [Theory]
    [InlineData("Ниязова М.Е.", "niyazova.me")]
    [InlineData("Хамит А.Ү.", "khamit.au")]
    [InlineData("Байтукешов Т.С.", "baytukeshov.ts")]
    [InlineData("Халель", "khalel")]
    public void Login_IsTransliterated(string name, string login) => Assert.Equal(login, ScheduleImportService.Login(name));

    [Fact]
    public async Task ScheduleImport_CreatesTeachersGroupsAndLessons_Idempotently()
    {
        await AuthenticateAsync("teacher", "teacher123");
        var file = CollegeSchedule();
        var preview = await PostFileAsync<ScheduleImportPreviewDto>("/api/schedule/import-xlsx/preview", file);
        Assert.Equal(new DateOnly(2026, 8, 31), preview.SemesterStart);
        Assert.Equal(2, preview.Groups.Count);
        Assert.Contains("Ниязова М.Е.", preview.Teachers);
        Assert.Single(preview.Warnings);

        var result = await PostFileAsync<ScheduleImportResultDto>("/api/schedule/import-xlsx?groups=23-ПО-103-4р-1", file);
        Assert.Equal(4, result.Added);
        Assert.Equal(0, result.Skipped);
        var niyazova = Assert.Single(result.Teachers, x => x.FullName == "Ниязова М.Е.");
        Assert.True(niyazova is { Created: true, Active: true, Login: "niyazova.me" });
        Assert.NotNull(niyazova.TemporaryPassword);
        var vacancy = Assert.Single(result.Teachers, x => x.FullName == CollegeScheduleParser.Vacancy);
        Assert.True(vacancy is { Active: false, TemporaryPassword: null });

        var again = await PostFileAsync<ScheduleImportResultDto>("/api/schedule/import-xlsx?groups=23-ПО-103-4р-1", file);
        Assert.Equal(0, again.Added);
        Assert.Equal(4, again.Skipped);
        Assert.All(again.Teachers, x => Assert.False(x.Created));

        // Новый преподаватель входит временным паролем и видит своё занятие в 6-ю неделю (понедельник 05.10.2026).
        using var teacherClient = factory.CreateClient();
        var token = await LoginAsync(teacherClient, "niyazova.me", niyazova.TemporaryPassword!);
        teacherClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token.AccessToken);
        var day = await teacherClient.GetFromJsonAsync<List<LessonDto>>("/api/schedule/day?date=2026-10-05", JsonOptions);
        Assert.Contains(day!, x => x.Slot is { Subject: "Искусственный интеллект OpenCV", Room: "711", GroupName: "23-ПО-103-4р-1" });
        Assert.DoesNotContain((await teacherClient.GetFromJsonAsync<List<LessonDto>>("/api/schedule/day?date=2026-10-19", JsonOptions))!,
            x => x.Slot.Subject == "Искусственный интеллект OpenCV");
        // Импорт создаёт учётные записи — неадмину запрещён.
        var forbidden = await teacherClient.PostAsync("/api/schedule/import-xlsx?groups=23-ПО-103-4р-1", new ByteArrayContent(file));
        Assert.Equal(HttpStatusCode.Forbidden, forbidden.StatusCode);
    }

    [Fact]
    public async Task ScheduleImport_ReusesTeacherAccountAcrossSpellingVariants()
    {
        await AuthenticateAsync("teacher", "teacher123");
        const string path = "/api/schedule/import-xlsx?groups=23-ПО-103-4р-1";
        var first = await PostFileAsync<ScheduleImportResultDto>(path, CollegeScheduleWithTeacher("Тайгара А."));
        var created = Assert.Single(first.Teachers);
        Assert.True(created.Created);

        foreach (var name in new[] { "Тайғара А.М.", "Тайгара" })
        {
            var again = await PostFileAsync<ScheduleImportResultDto>(path, CollegeScheduleWithTeacher(name));
            Assert.Equal((0, 1), (again.Added, again.Skipped));
            var reused = Assert.Single(again.Teachers);
            Assert.False(reused.Created);
            Assert.Equal((created.FullName, created.Login), (reused.FullName, reused.Login));
        }
    }

    private static byte[] CollegeScheduleWithTeacher(string name, string room = "811", string weeks = "5")
    {
        using var source = new MemoryStream(CollegeSchedule());
        using var workbook = new XLWorkbook(source);
        var sheet = workbook.Worksheet("4 курс");
        sheet.Cell(7, 5).Value = $"Графика · {room} · нед. {weeks}";
        sheet.Cell(8, 5).Value = name;
        sheet.Cell(9, 5).Value = "";
        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        return stream.ToArray();
    }

    private static byte[] CollegeScheduleUpdated()
    {
        using var source = new MemoryStream(CollegeSchedule());
        using var workbook = new XLWorkbook(source);
        var sheet = workbook.Worksheet("4 курс");
        sheet.Cell(7, 5).Value = "Искусственный интеллект OpenCV · 712 · нед. 6-7";
        sheet.Cell(8, 5).Value = "Ниязова М.Е. · 712";
        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        return stream.ToArray();
    }

    [Fact]
    public async Task ScheduleImport_ReplacesOutdatedLessonsOfSelectedGroups()
    {
        await AuthenticateAsync("teacher", "teacher123");
        var first = await PostFileAsync<ScheduleImportResultDto>("/api/schedule/import-xlsx?groups=23-ПО-103-4р-1&groups=23-ПО-103-4р-2", CollegeSchedule());
        Assert.Equal(6, first.Added);
        using var scope = factory.Services.CreateScope();
        var db = scope.ServiceProvider.GetRequiredService<AppDbContext>();
        var secondGroupId = await db.Groups.Where(x => x.Name == "23-ПО-103-4р-2").Select(x => x.Id).SingleAsync();
        var secondGroupSlots = await db.ScheduleSlots.CountAsync(x => x.GroupId == secondGroupId);

        var updated = await PostFileAsync<ScheduleImportResultDto>("/api/schedule/import-xlsx?groups=23-ПО-103-4р-1", CollegeScheduleUpdated());
        Assert.Equal((1, 2, 2), (updated.Added, updated.Skipped, updated.Removed));
        db.ChangeTracker.Clear();
        var slots = await db.ScheduleSlots.Where(x => x.Group.Name == "23-ПО-103-4р-1").ToListAsync();
        Assert.Equal(3, slots.Count);
        Assert.DoesNotContain(slots, x => x.Room == "711" || x.Subject == "База данных PL SQL/MONGODB");
        Assert.Contains(slots, x => x is { Subject: "Искусственный интеллект OpenCV", Room: "712", Imported: true });
        Assert.Equal(secondGroupSlots, await db.ScheduleSlots.CountAsync(x => x.GroupId == secondGroupId));

        var again = await PostFileAsync<ScheduleImportResultDto>("/api/schedule/import-xlsx?groups=23-ПО-103-4р-1", CollegeScheduleUpdated());
        Assert.Equal((0, 3, 0), (again.Added, again.Skipped, again.Removed));
    }

    [Fact]
    public async Task ScheduleImport_TrimsOutdatedLessonWithAttendance()
    {
        await AuthenticateAsync("teacher", "teacher123");
        const string path = "/api/schedule/import-xlsx?groups=23-ПО-103-4р-1";
        await PostFileAsync<ScheduleImportResultDto>(path, CollegeScheduleWithTeacher("Тайгара А.", "811", "5-18"));
        using var scope = factory.Services.CreateScope();
        var db = scope.ServiceProvider.GetRequiredService<AppDbContext>();
        var old = await db.ScheduleSlots.SingleAsync(x => x.Group.Name == "23-ПО-103-4р-1");
        var student = new Student { GroupId = old.GroupId, FullName = "Иван Иванов" };
        db.Students.Add(student);
        await db.SaveChangesAsync();
        var markedDate = old.ValidFrom;
        db.AttendanceMarks.Add(new AttendanceMark { SlotId = old.Id, StudentId = student.Id, Date = markedDate,
            Status = AttendanceStatus.Present, Method = MarkMethod.Manual, MarkedAt = DateTime.Now });
        await db.SaveChangesAsync();

        var updated = await PostFileAsync<ScheduleImportResultDto>(path, CollegeScheduleWithTeacher("Тайгара А.", "812", "5-18"));
        Assert.Equal((1, 1), (updated.Added, updated.Removed));
        db.ChangeTracker.Clear();
        var historic = await db.ScheduleSlots.SingleAsync(x => x.Id == old.Id);
        Assert.Equal(DateOnly.FromDateTime(DateTime.Now).AddDays(-1), historic.ValidTo);
        Assert.True(historic.Imported);
        Assert.True(await db.AttendanceMarks.AnyAsync(x => x.SlotId == old.Id && x.Date == markedDate));
    }

    [Fact]
    public async Task ScheduleImport_KeepsManualLesson()
    {
        await AuthenticateAsync("teacher", "teacher123");
        const string path = "/api/schedule/import-xlsx?groups=23-ПО-103-4р-1";
        await PostFileAsync<ScheduleImportResultDto>(path, CollegeScheduleWithTeacher("Тайгара А."));
        using var scope = factory.Services.CreateScope();
        var db = scope.ServiceProvider.GetRequiredService<AppDbContext>();
        var old = await db.ScheduleSlots.SingleAsync(x => x.Group.Name == "23-ПО-103-4р-1");
        var manual = await scope.ServiceProvider.GetRequiredService<ScheduleService>().AddSlotAsync(new ScheduleSlot
        {
            TeacherId = old.TeacherId, GroupId = old.GroupId, Subject = "Ручная пара", DayOfWeek = DayOfWeek.Tuesday,
            Start = new TimeOnly(11, 0), End = new TimeOnly(12, 20), Parity = WeekParity.Every, ValidFrom = old.ValidFrom
        });
        Assert.False(manual.Imported);

        var updated = await PostFileAsync<ScheduleImportResultDto>(path, CollegeScheduleWithTeacher("Тайгара А.", "812"));
        Assert.Equal(1, updated.Removed);
        db.ChangeTracker.Clear();
        var persisted = await db.ScheduleSlots.SingleAsync(x => x.Id == manual.Id);
        Assert.False(persisted.Imported);
        Assert.Equal("Ручная пара", persisted.Subject);
    }

    [Fact]
    public async Task ScheduleImport_EditingImportedLessonMakesItManual()
    {
        await AuthenticateAsync("teacher", "teacher123");
        const string path = "/api/schedule/import-xlsx?groups=23-ПО-103-4р-1";
        var first = await PostFileAsync<ScheduleImportResultDto>(path, CollegeScheduleWithTeacher("Тайгара А."));
        var teacher = Assert.Single(first.Teachers);
        using var scope = factory.Services.CreateScope();
        var db = scope.ServiceProvider.GetRequiredService<AppDbContext>();
        var slot = await db.ScheduleSlots.AsNoTracking().SingleAsync(x => x.Group.Name == "23-ПО-103-4р-1");
        using var teacherClient = factory.CreateClient();
        var token = await LoginAsync(teacherClient, teacher.Login, teacher.TemporaryPassword!);
        teacherClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token.AccessToken);
        var request = new SaveSlotRequest(slot.GroupId, null, slot.Subject, "812", slot.DayOfWeek,
            slot.Start, slot.End, slot.Parity, slot.OneOffDate);
        var response = await teacherClient.PutAsJsonAsync($"/api/schedule/slots/{slot.Id}", request, JsonOptions);
        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        db.ChangeTracker.Clear();
        Assert.False(await db.ScheduleSlots.Where(x => x.Id == slot.Id).Select(x => x.Imported).SingleAsync());

        var again = await PostFileAsync<ScheduleImportResultDto>(path, CollegeScheduleWithTeacher("Тайгара А."));
        Assert.Equal((1, 0), (again.Added, again.Removed));
        db.ChangeTracker.Clear();
        var edited = await db.ScheduleSlots.SingleAsync(x => x.Id == slot.Id);
        Assert.Equal("812", edited.Room);
        Assert.False(edited.Imported);
    }

    [Fact]
    public async Task ScheduleImport_RejectsUnknownGroupsAndBrokenFiles()
    {
        await AuthenticateAsync("teacher", "teacher123");
        var unknown = await client.PostAsync("/api/schedule/import-xlsx?groups=99-XX-000-1", new ByteArrayContent(CollegeSchedule()));
        Assert.Equal(HttpStatusCode.BadRequest, unknown.StatusCode);
        var none = await client.PostAsync("/api/schedule/import-xlsx", new ByteArrayContent(CollegeSchedule()));
        Assert.Equal(HttpStatusCode.BadRequest, none.StatusCode);
        var broken = await client.PostAsync("/api/schedule/import-xlsx/preview", new ByteArrayContent("не excel"u8.ToArray()));
        Assert.Equal(HttpStatusCode.BadRequest, broken.StatusCode);
    }

    [Fact]
    public async Task StudentImport_ReadsJournalWorkbook()
    {
        await AuthenticateAsync("teacher", "teacher123");
        using var workbook = new XLWorkbook();
        foreach (var (group, names) in new[] { ("99-ИМ-001-1", new[] { "Абдиев Артур ", " Ковалев  Тимур" }), ("99-ИМ-001-2", new[] { "Ли Кирилл" }) })
        {
            var sheet = workbook.Worksheets.Add(group);
            sheet.Cell(2, 1).Value = "#"; sheet.Cell(2, 2).Value = "ФИО"; sheet.Cell(2, 3).Value = 16;
            for (var i = 0; i < names.Length; i++) { sheet.Cell(3 + i, 1).Value = i + 1; sheet.Cell(3 + i, 2).Value = names[i]; }
            sheet.Cell(4 + names.Length, 2).Value = "куратор:Рахимкужаева Алия Рахимберлиевна";
        }
        workbook.Worksheets.Add("Пустой").Cell(1, 1).Value = "заметки";
        using var stream = new MemoryStream();
        workbook.SaveAs(stream);

        var result = await PostFileAsync<ImportResultDto>("/api/students/import", stream.ToArray());
        Assert.Equal(3, result.Added);
        Assert.Equal(0, result.Skipped);
        Assert.Single(result.Errors, x => x.Contains("Пустой"));
        var groups = await client.GetFromJsonAsync<List<GroupDto>>("/api/groups", JsonOptions);
        var first = Assert.Single(groups!, x => x.Name == "99-ИМ-001-1");
        var students = await client.GetFromJsonAsync<List<StudentDto>>($"/api/groups/{first.Id}/students", JsonOptions);
        Assert.Equal(["Абдиев Артур", "Ковалев Тимур"], students!.Select(x => x.FullName).Order());

        var again = await PostFileAsync<ImportResultDto>("/api/students/import", stream.ToArray());
        Assert.Equal((0, 3), (again.Added, again.Skipped));
        // CSV по-прежнему поддерживается тем же маршрутом.
        var csv = await PostFileAsync<ImportResultDto>("/api/students/import", "ФИО;Группа\r\nНовиков Пётр;99-ИМ-001-2\r\n"u8.ToArray());
        Assert.Equal(1, csv.Added);
    }

    private async Task<T> PostFileAsync<T>(string path, byte[] file)
    {
        var response = await client.PostAsync(path, new ByteArrayContent(file));
        Assert.True(response.StatusCode == HttpStatusCode.OK, $"{path}: {(int)response.StatusCode} {await response.Content.ReadAsStringAsync()}");
        return (await response.Content.ReadFromJsonAsync<T>(JsonOptions))!;
    }

    private async Task AuthenticateAsync(string login, string password)
    {
        var token = await LoginAsync(client, login, password);
        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token.AccessToken);
    }

    private static async Task<TokenResponse> LoginAsync(HttpClient target, string login, string password)
    {
        var response = await target.PostAsJsonAsync("/api/auth/login", new LoginRequest(login, password), JsonOptions);
        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        return (await response.Content.ReadFromJsonAsync<TokenResponse>(JsonOptions))!;
    }

    public void Dispose()
    {
        client.Dispose();
        factory.Dispose();
        Directory.Delete(temporaryDirectory, recursive: true);
    }

    private sealed class ImportFactory(string directory) : WebApplicationFactory<Program>
    {
        protected override void ConfigureWebHost(IWebHostBuilder builder)
        {
            builder.UseContentRoot(ServerRoot());
            builder.UseSetting("ConnectionStrings:Default", $"Data Source={Path.Combine(directory, "test.db")};Pooling=False");
            builder.UseSetting("Demo:Enabled", "true");
        }

        private static string ServerRoot()
        {
            for (var current = new DirectoryInfo(AppContext.BaseDirectory); current is not null; current = current.Parent)
            {
                var candidate = Path.Combine(current.FullName, "BiStudent.Server");
                if (File.Exists(Path.Combine(candidate, "seed", "students.demo.csv"))) return candidate;
            }
            throw new DirectoryNotFoundException("BiStudent.Server/seed/students.demo.csv");
        }
    }
}
