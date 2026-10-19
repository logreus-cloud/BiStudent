using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace BiStudent.Server.Data.Migrations
{
    /// <inheritdoc />
    public partial class ScheduleSlotImported : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<bool>(
                name: "Imported",
                table: "ScheduleSlots",
                type: "INTEGER",
                nullable: false,
                defaultValue: false);

            // Пары, созданные до появления флага, пришли из импорта расписания.
            migrationBuilder.Sql("UPDATE \"ScheduleSlots\" SET \"Imported\" = 1");
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "Imported",
                table: "ScheduleSlots");
        }
    }
}
