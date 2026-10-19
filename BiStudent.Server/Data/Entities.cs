using BiStudent.Contracts;

namespace BiStudent.Server.Data;

public sealed class Teacher
{
    public int Id { get; set; }
    public string Login { get; set; } = "";
    public string FullName { get; set; } = "";
    public string PasswordHash { get; set; } = "";
    public bool IsAdmin { get; set; }
    public bool IsActive { get; set; } = true;
    public bool MustChangePassword { get; set; }
    public string SecurityStamp { get; set; } = Guid.NewGuid().ToString("N");
    public DateTime CreatedAt { get; set; } = DateTime.Now;
    public DateTime? LastLoginAt { get; set; }
}

public sealed class Group
{
    public int Id { get; set; }
    public string Name { get; set; } = "";
    public List<Student> Students { get; set; } = [];
}

public sealed class Student
{
    public int Id { get; set; }
    public string FullName { get; set; } = "";
    public int GroupId { get; set; }
    public Group Group { get; set; } = null!;
    public DateTime CreatedAt { get; set; } = DateTime.Now;
    public List<FaceSample> Samples { get; set; } = [];
    public List<GroupMembership> Memberships { get; set; } = [];
}

public sealed class GroupMembership
{
    public int Id { get; set; }
    public int StudentId { get; set; }
    public Student Student { get; set; } = null!;
    public int GroupId { get; set; }
    public Group Group { get; set; } = null!;
    public DateOnly From { get; set; }
    public DateOnly? To { get; set; }
}

public sealed class FaceSample
{
    public int Id { get; set; }
    public int StudentId { get; set; }
    public Student Student { get; set; } = null!;
    public byte[] ProtectedImage { get; set; } = [];
    public DateTime CreatedAt { get; set; } = DateTime.Now;
}

public sealed class ScheduleSlot
{
    public int Id { get; set; }
    public int TeacherId { get; set; }
    public Teacher Teacher { get; set; } = null!;
    public int GroupId { get; set; }
    public Group Group { get; set; } = null!;
    public string Subject { get; set; } = "";
    public string? Room { get; set; }
    public DayOfWeek DayOfWeek { get; set; }
    public TimeOnly Start { get; set; }
    public TimeOnly End { get; set; }
    public WeekParity Parity { get; set; }
    public DateOnly ValidFrom { get; set; }
    public DateOnly? ValidTo { get; set; }
    public DateOnly? OneOffDate { get; set; }
    public bool Imported { get; set; }
    public List<SlotCancellation> Cancellations { get; set; } = [];
}

public sealed class SlotCancellation
{
    public int Id { get; set; }
    public int SlotId { get; set; }
    public ScheduleSlot Slot { get; set; } = null!;
    public DateOnly Date { get; set; }
    public string Reason { get; set; } = "";
}

public sealed class AttendanceMark
{
    public int Id { get; set; }
    public int SlotId { get; set; }
    public ScheduleSlot Slot { get; set; } = null!;
    public DateOnly Date { get; set; }
    public int StudentId { get; set; }
    public Student Student { get; set; } = null!;
    public AttendanceStatus Status { get; set; }
    public MarkMethod Method { get; set; }
    public DateTime MarkedAt { get; set; }
    public ExcuseReason? ExcuseReason { get; set; }
    public string? Comment { get; set; }
    public bool HasDocument { get; set; }
}

public sealed class LessonGrade
{
    public int Id { get; set; }
    public int SlotId { get; set; }
    public ScheduleSlot Slot { get; set; } = null!;
    public DateOnly Date { get; set; }
    public int StudentId { get; set; }
    public Student Student { get; set; } = null!;
    public int Score { get; set; }
    public DateTime UpdatedAt { get; set; } = DateTime.Now;
}

public sealed class Assessment
{
    public int Id { get; set; }
    public int TeacherId { get; set; }
    public Teacher Teacher { get; set; } = null!;
    public int GroupId { get; set; }
    public Group Group { get; set; } = null!;
    public string Subject { get; set; } = "";
    public string Title { get; set; } = "";
    public AssessmentKind Kind { get; set; }
    public DateOnly Date { get; set; }
    public List<AssessmentGrade> Grades { get; set; } = [];
}

public sealed class AssessmentGrade
{
    public int Id { get; set; }
    public int AssessmentId { get; set; }
    public Assessment Assessment { get; set; } = null!;
    public int StudentId { get; set; }
    public Student Student { get; set; } = null!;
    public int Score { get; set; }
    public DateTime UpdatedAt { get; set; } = DateTime.Now;
}

public sealed class FinalGrade
{
    public int Id { get; set; }
    public int TeacherId { get; set; }
    public Teacher Teacher { get; set; } = null!;
    public int GroupId { get; set; }
    public Group Group { get; set; } = null!;
    public string Subject { get; set; } = "";
    public int StudentId { get; set; }
    public Student Student { get; set; } = null!;
    public int Score { get; set; }
    public DateTime UpdatedAt { get; set; } = DateTime.Now;
}

public sealed class EnrollmentCode
{
    public int Id { get; set; }
    public int GroupId { get; set; }
    public Group Group { get; set; } = null!;
    public string Code { get; set; } = "";
    public int CreatedByTeacherId { get; set; }
    public Teacher CreatedBy { get; set; } = null!;
    public DateTime CreatedAt { get; set; } = DateTime.Now;
    public DateTime ExpiresAt { get; set; }
    public DateTime? RevokedAt { get; set; }
}

public sealed class FaceSubmission
{
    public int Id { get; set; }
    public int StudentId { get; set; }
    public Student Student { get; set; } = null!;
    public int? EnrollmentCodeId { get; set; }
    public EnrollmentCode? EnrollmentCode { get; set; }
    public DateTime CreatedAt { get; set; } = DateTime.Now;
    public FaceSubmissionStatus Status { get; set; }
    public string? SimilarTo { get; set; }
    public double SimilarScore { get; set; }
    public DateTime? ReviewedAt { get; set; }
    public int? ReviewedByTeacherId { get; set; }
    public string? RejectReason { get; set; }
    public List<FaceSubmissionImage> Images { get; set; } = [];
}

public sealed class FaceSubmissionImage
{
    public int Id { get; set; }
    public int SubmissionId { get; set; }
    public FaceSubmission Submission { get; set; } = null!;
    // Выровненный кроп лица, зашифрованный тем же протектором, что и FaceSample.
    public byte[] ProtectedImage { get; set; } = [];
}

public sealed class AppSetting
{
    public string Key { get; set; } = "";
    public string Value { get; set; } = "";
}
