using System.Globalization;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

namespace BiStudent.Server.Services;

/// Одно занятие из файла: либо еженедельное в диапазоне недель (FromWeek–ToWeek), либо разовое (FromWeek == ToWeek, Once).
public sealed record ParsedLesson(string Sheet, string Group, DayOfWeek Day, TimeOnly Start, TimeOnly End,
    string Subject, string? Room, string Teacher, int FromWeek, int ToWeek, bool Once);

public sealed record ParsedSchedule(DateOnly SemesterStart, List<ParsedLesson> Lessons, List<(string Sheet, string Group)> Groups, List<string> Warnings);

/// Разбор расписания колледжа в формате «блок из пяти колонок на группу»:
/// день · время звонка · № пары · аудитория · группа. В ячейке группы — строки «[ТИП] Предмет · аудитория · нед. 2-4, 8 (разово)»,
/// строкой ниже — преподаватели в том же порядке («Фамилия И.О. · аудитория»).
public static partial class CollegeScheduleParser
{
    public const string Vacancy = "Вакансия";

    public static ParsedSchedule Parse(Stream xlsx, DateOnly fallbackSemesterStart)
    {
        using var workbook = new XLWorkbook(xlsx);
        var lessons = new List<ParsedLesson>();
        var groups = new List<(string, string)>();
        var warnings = new List<string>();
        var teacherNames = new List<string>();
        DateOnly? semester = null;
        foreach (var sheet in workbook.Worksheets)
        {
            var used = sheet.RangeUsed();
            if (used is null) continue;
            var lastRow = used.LastRow().RowNumber();
            var lastColumn = used.LastColumn().ColumnNumber();
            semester ??= FindSemesterStart(sheet, Math.Min(lastRow, 10), lastColumn);
            var header = FindHeaderRow(sheet, Math.Min(lastRow, 30), lastColumn);
            if (header is not int headerRow) continue;
            for (var column = 5; column <= lastColumn; column++)
            {
                var name = Text(sheet.Cell(headerRow, column));
                if (!GroupName().IsMatch(name)) continue;
                groups.Add((sheet.Name, name));
                ParseGroup(sheet, name, headerRow, column, lastRow, lessons, warnings, teacherNames);
            }
        }
        return new ParsedSchedule(semester ?? fallbackSemesterStart, UnifyTeachers(lessons, teacherNames), groups, warnings);
    }

    private static void ParseGroup(IXLWorksheet sheet, string group, int headerRow, int column, int lastRow,
        List<ParsedLesson> lessons, List<string> warnings, List<string> teacherNames)
    {
        int dayColumn = column - 4, timeColumn = column - 3;
        for (var row = headerRow + 1; row <= lastRow; row++)
        {
            var timeCell = sheet.Cell(row, timeColumn);
            var merged = timeCell.MergedRange();
            // Пара занимает две строки (предмет и преподаватель); берём первую строку объединённой ячейки времени.
            if (merged is not null && merged.FirstRow().RowNumber() != row) continue;
            var time = Text(merged?.FirstCell() ?? timeCell);
            var times = TimePattern().Matches(time);
            if (times.Count < 2) continue;
            var dayCell = sheet.Cell(row, dayColumn);
            var day = ParseDay(Text(dayCell.MergedRange()?.FirstCell() ?? dayCell));
            if (day is not DayOfWeek weekday) continue;
            var start = TimeOnly.ParseExact(times[0].Value, ["H:mm", "HH:mm"], CultureInfo.InvariantCulture);
            var end = TimeOnly.ParseExact(times[1].Value, ["H:mm", "HH:mm"], CultureInfo.InvariantCulture);
            var subjectCell = sheet.Cell(row, column);
            var subjectText = Text(subjectCell);
            if (subjectText.Length == 0) continue;
            var teacherRow = merged?.LastRow().RowNumber() is int last && last > row ? last : row + 1;
            var teachers = ExtractTeachers(Text(sheet.Cell(teacherRow, column)));
            teacherNames.AddRange(teachers);
            var entries = SplitEntries(subjectText);
            for (var i = 0; i < entries.Count; i++)
            {
                var teacher = teachers.Count == 0 ? Vacancy : teachers[Math.Min(i, teachers.Count - 1)];
                if (!TryParseEntry(entries[i], out var subject, out var room, out var weeks))
                {
                    warnings.Add($"{sheet.Name}, {group}, {DayName(weekday)} {start:HH\\:mm}: не удалось разобрать «{entries[i]}».");
                    continue;
                }
                foreach (var (from, to, once) in weeks)
                    lessons.Add(new ParsedLesson(sheet.Name, group, weekday, start, end, subject, room, teacher, from, to, once));
            }
        }
    }

    /// Строка ячейки может содержать несколько записей: по строкам, через два и более пробела или перед новым «[ТИП]».
    public static List<string> SplitEntries(string text)
    {
        var result = new List<string>();
        foreach (var line in text.Split('\n'))
        {
            var chunks = EntrySeparator().Split(line).Select(x => x.Trim()).Where(x => x.Length > 0).ToList();
            for (var i = 0; i < chunks.Count; i++)
            {
                if (chunks[i].Equals("СРОП", StringComparison.OrdinalIgnoreCase) &&
                    i + 1 < chunks.Count && chunks[i + 1].StartsWith('[')) continue;
                // Двойной пробел внутри текста или списка недель не начинает новую запись.
                if (i > 0 && result.Count > 0 && chunks[i][0] != '[' &&
                    (char.IsLower(chunks[i][0]) || char.IsDigit(chunks[i][0]) || chunks[i][0] == ')' ||
                     result[^1].EndsWith(',') || result[^1].EndsWith("нед.", StringComparison.OrdinalIgnoreCase)))
                    result[^1] += " " + chunks[i];
                else result.Add(chunks[i]);
            }
        }
        return result;
    }

    public static bool TryParseEntry(string entry, out string subject, out string? room, out List<(int From, int To, bool Once)> weeks)
    {
        subject = ""; room = null; weeks = [];
        var tag = TagPattern().Match(entry);
        var kind = tag.Success ? tag.Groups[1].Value.Trim().ToUpperInvariant() : "";
        var rest = tag.Success ? entry[tag.Length..].Trim() : entry.Trim();
        var trailingRoom = WeeksRoomSuffix().Match(rest);
        string? roomAfterWeeks = null;
        if (trailingRoom.Success)
        {
            roomAfterWeeks = trailingRoom.Groups["room"].Value;
            rest = rest[..trailingRoom.Index].TrimEnd();
        }
        var weeksMatch = WeeksPattern().Match(rest);
        if (!weeksMatch.Success) weeksMatch = TrailingWeeksPattern().Match(rest);
        if (!weeksMatch.Success) return false;
        weeks = ParseWeeks(weeksMatch.Groups["weeks"].Value);
        if (weeks.Count == 0) return false;
        var head = rest[..weeksMatch.Index].Trim().TrimEnd('·', '-', ',', ' ');
        var parts = head.Split('·').Select(x => x.Trim()).Where(x => x.Length > 0).ToList();
        if (parts.Count == 1 && DashRoom().Match(parts[0]) is { Success: true } dashed)
            parts = [dashed.Groups["subject"].Value.Trim(), dashed.Groups["room"].Value.Trim()];
        parts.RemoveAll(x => x.StartsWith("с ", StringComparison.OrdinalIgnoreCase) &&
            GroupName().IsMatch(x[2..].Split(',')[0].Trim()));
        if (kind == "СРОП" && parts.Count > 0 && parts[0].Equals("СРОП", StringComparison.OrdinalIgnoreCase)) parts.RemoveAt(0);
        if (parts.Count == 0) return false;
        subject = Regex.Replace(parts[0], @"\s+", " ");
        room = roomAfterWeeks ?? (parts.Count > 1 ? Room(parts[1]) : null);
        return subject.Length > 0;
    }

    public static List<(int From, int To, bool Once)> ParseWeeks(string text)
    {
        var result = new List<(int, int, bool)>();
        var explicitOnce = new List<bool>();
        foreach (var raw in text.Split(','))
        {
            var token = raw.Trim().Replace('–', '-').Replace('—', '-');
            if (token.Length == 0) continue;
            var once = token.Contains("разово", StringComparison.OrdinalIgnoreCase);
            var numbers = new List<int>();
            foreach (Match match in Regex.Matches(token, @"\d+"))
            {
                if (!int.TryParse(match.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var n))
                {
                    numbers.Clear();
                    break;
                }
                numbers.Add(n);
            }
            if (numbers.Count == 0 || numbers.Any(x => x is < 1 or > 60)) continue;
            var from = numbers[0];
            var to = numbers.Count > 1 && token.Contains('-') ? numbers[1] : from;
            if (to < from) (from, to) = (to, from);
            // Подряд идущие недели без «разово» («нед.1,2,3,4») — одна еженедельная пара, а не набор разовых.
            if (!once && result.Count > 0 && result[^1] is var (lastFrom, lastTo, _) && lastTo + 1 == from && !explicitOnce[^1])
                result[^1] = (lastFrom, to, false);
            else
            {
                result.Add((from, to, once || from == to));
                explicitOnce.Add(once);
            }
        }
        return result;
    }

    public static List<string> ExtractTeachers(string text)
    {
        var result = new List<string>();
        text = Regex.Replace(text, @"Спортзал\s*\(поток\s*\d+\)", " ", RegexOptions.IgnoreCase);
        foreach (Match match in TeacherPattern().Matches(text))
        {
            if (match.Groups["vacancy"].Success) { result.Add(Vacancy); continue; }
            if (match.Groups["mentor"].Success) { result.Add(match.Value); continue; }
            var surname = match.Groups["surname"].Value;
            if (surname.Equals("уточнить", StringComparison.OrdinalIgnoreCase) ||
                surname.Equals("Спортзал", StringComparison.OrdinalIgnoreCase)) continue;
            var initials = string.Concat(match.Groups["initials"].Value.Where(char.IsLetter).Select(c => $"{c}."));
            result.Add(initials.Length == 0 ? surname : $"{surname} {initials}");
        }
        return result;
    }

    public static string TeacherName(string line) => ExtractTeachers(line).FirstOrDefault() ?? Vacancy;

    public static List<string> UnifyTeacherNames(IReadOnlyList<string> names)
    {
        var identities = names.Select(TeacherIdentity).ToList();
        var result = new List<string>(names.Count);
        for (var i = 0; i < names.Count; i++)
        {
            if (names[i] == Vacancy || identities[i] is not { } identity) { result.Add(names[i]); continue; }
            var candidates = Enumerable.Range(0, names.Count).Where(j =>
                identities[j] is { } other && other.Surname == identity.Surname &&
                other.Initials.StartsWith(identity.Initials, StringComparison.Ordinal)).ToList();
            var length = candidates.Max(j => identities[j]!.Value.Initials.Length);
            var matches = candidates.Where(j => identities[j]!.Value.Initials.Length == length).ToList();
            if (matches.Select(j => identities[j]!.Value.Initials).Distinct().Count() != 1)
                matches = candidates.Where(j => identities[j]!.Value.Initials == identity.Initials).ToList();
            result.Add(matches.GroupBy(j => names[j], StringComparer.OrdinalIgnoreCase)
                .OrderByDescending(g => g.Count()).ThenByDescending(g => g.Key.Count(c => c == '.')).ThenBy(g => g.Min())
                .Select(g => names[g.First()]).First());
        }
        return result;
    }

    public static bool TeacherNamesMatch(string name, string existingName)
    {
        if (name.Equals(existingName, StringComparison.OrdinalIgnoreCase)) return true;
        if (name == Vacancy || existingName == Vacancy) return false;
        var first = TeacherIdentity(name);
        var second = TeacherIdentity(existingName);
        return first is { } a && second is { } b && a.Surname == b.Surname &&
            (a.Initials.StartsWith(b.Initials, StringComparison.Ordinal) ||
             b.Initials.StartsWith(a.Initials, StringComparison.Ordinal));
    }

    private static List<ParsedLesson> UnifyTeachers(List<ParsedLesson> lessons, List<string> teacherNames)
    {
        var unified = UnifyTeacherNames(teacherNames);
        var canonical = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < teacherNames.Count; i++) canonical[teacherNames[i]] = unified[i];
        return lessons.Select(lesson => canonical.TryGetValue(lesson.Teacher, out var name)
            ? lesson with { Teacher = name } : lesson).ToList();
    }

    private static (string Surname, string Initials)? TeacherIdentity(string name)
    {
        var match = TeacherIdentityPattern().Match(name);
        if (!match.Success) return null;
        var initials = new string(match.Groups["initials"].Value.Where(char.IsLetter).ToArray());
        return (Fold(match.Groups["surname"].Value), Fold(initials));
    }

    private static string Fold(string value) => new(value.ToLowerInvariant().Select(c => c switch
    {
        'ғ' => 'г', 'қ' => 'к', 'ү' or 'ұ' => 'у', 'ә' => 'а', 'ө' => 'о',
        'ң' => 'н', 'һ' => 'х', 'і' => 'и', _ => c
    }).ToArray());

    private static string? Room(string value)
    {
        var room = value.Trim().TrimEnd('-').Trim();
        if (room.StartsWith("уточн", StringComparison.OrdinalIgnoreCase))
        {
            var numbered = UncertainRoom().Match(room);
            return numbered.Success ? numbered.Groups["room"].Value : null;
        }
        return room.Length == 0 ? null : room;
    }

    private static int? FindHeaderRow(IXLWorksheet sheet, int lastRow, int lastColumn)
    {
        for (var row = 1; row <= lastRow; row++)
            for (var column = 5; column <= lastColumn; column++)
                if (GroupName().IsMatch(Text(sheet.Cell(row, column))) &&
                    Text(sheet.Cell(row, column - 4)).Contains("Дни", StringComparison.OrdinalIgnoreCase))
                    return row;
        return null;
    }

    private static DateOnly? FindSemesterStart(IXLWorksheet sheet, int lastRow, int lastColumn)
    {
        for (var row = 1; row <= lastRow; row++)
            for (var column = 1; column <= Math.Min(lastColumn, 5); column++)
                if (SemesterPattern().Match(Text(sheet.Cell(row, column))) is { Success: true } match &&
                    DateOnly.TryParseExact(match.Groups[1].Value, "dd.MM.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var date))
                    return date.AddDays(-(((int)date.DayOfWeek + 6) % 7));
        return null;
    }

    private static DayOfWeek? ParseDay(string text)
    {
        var value = text.ToLowerInvariant();
        if (value.Contains("понедельник")) return DayOfWeek.Monday;
        if (value.Contains("вторник")) return DayOfWeek.Tuesday;
        if (value.Contains("среда")) return DayOfWeek.Wednesday;
        if (value.Contains("четверг")) return DayOfWeek.Thursday;
        if (value.Contains("пятница")) return DayOfWeek.Friday;
        if (value.Contains("суббота")) return DayOfWeek.Saturday;
        if (value.Contains("воскресенье")) return DayOfWeek.Sunday;
        return null;
    }

    public static string DayName(DayOfWeek day) => CultureInfo.GetCultureInfo("ru-RU").DateTimeFormat.GetAbbreviatedDayName(day);

    private static string Text(IXLCell cell) => cell.GetFormattedString().Replace("\r", "").Trim();

    [GeneratedRegex(@"^\d{2}-[А-ЯЁA-Z]{1,4}-\d{2,4}-\S+$")]
    private static partial Regex GroupName();

    [GeneratedRegex(@"\b\d{1,2}:\d{2}\b")]
    private static partial Regex TimePattern();

    [GeneratedRegex(@"^\[([^\]]+)\]")]
    private static partial Regex TagPattern();

    [GeneratedRegex(@"нед\.?\s*,?\s*(?<weeks>\d[\d\s,\-–—()а-яё]*)$", RegexOptions.IgnoreCase)]
    private static partial Regex WeeksPattern();

    [GeneratedRegex(@"\s-\s*(?<weeks>\d+(?:\s*-\s*\d+)?(?:\s*,\s*\d+(?:\s*-\s*\d+)?)*)\s*,?\s*$")]
    private static partial Regex TrailingWeeksPattern();

    [GeneratedRegex(@"^(?<subject>.+?)\s+-\s*(?<room>[^\s-][^-]*)$")]
    private static partial Regex DashRoom();

    [GeneratedRegex(@"\s{2,}|(?<=\S)\s+(?=\[[А-ЯЁA-Z]+\])")]
    private static partial Regex EntrySeparator();

    [GeneratedRegex(@"(?<=\d)\s+-\s*(?<room>\d{3,4})$")]
    private static partial Regex WeeksRoomSuffix();

    [GeneratedRegex(@"-(?<room>\d+)$")]
    private static partial Regex UncertainRoom();

    [GeneratedRegex(@"(?<vacancy>(?i:вакансия)(?:\s*\([^)]*\))?)|(?<mentor>Эрик\s*\(ментор\))|(?<![А-ЯЁа-яёӘҒҚҢӨҰҮҺІәғқңөұүһі])(?<surname>[А-ЯЁӘҒҚҢӨҰҮҺІ][а-яёәғқңөұүһі]+)(?:\s+(?<initials>[А-ЯЁӘҒҚҢӨҰҮҺІ](?:\.?[А-ЯЁӘҒҚҢӨҰҮҺІ])?\.?)(?![А-ЯЁа-яёӘҒҚҢӨҰҮҺІәғқңөұүһі\d]))?")]
    private static partial Regex TeacherPattern();

    [GeneratedRegex(@"^(?<surname>[А-ЯЁӘҒҚҢӨҰҮҺІ][а-яёәғқңөұүһі]+)(?:\s+(?<initials>[А-ЯЁӘҒҚҢӨҰҮҺІ]\.?(?:[А-ЯЁӘҒҚҢӨҰҮҺІ]\.?)?))?$")]
    private static partial Regex TeacherIdentityPattern();

    [GeneratedRegex(@"недели\s*\d+\s*-\s*\d+\s*,\s*(\d{2}\.\d{2}\.\d{4})", RegexOptions.IgnoreCase)]
    private static partial Regex SemesterPattern();
}
