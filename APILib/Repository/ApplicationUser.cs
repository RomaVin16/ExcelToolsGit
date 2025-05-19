using System;
using System.ComponentModel.DataAnnotations.Schema;
using Microsoft.AspNetCore.Identity;

namespace APILib.Repository
{
    [Table("aspnetusers")]
    public class ApplicationUser : IdentityUser
    {
        [Column("id")]
        public override string Id { get; set; }

        [Column("username")]
        public override string UserName { get; set; }

        [Column("normalizedusername")]
        public override string NormalizedUserName { get; set; }

        [Column("email")]
        public override string Email { get; set; }

        [Column("normalizedemail")]
        public override string NormalizedEmail { get; set; }

        [Column("emailconfirmed")]
        public override bool EmailConfirmed { get; set; }

        [Column("passwordhash")]
        public override string PasswordHash { get; set; }

        [Column("securitystamp")]
        public override string SecurityStamp { get; set; }

        [Column("concurrencystamp")]
        public override string ConcurrencyStamp { get; set; }

        [Column("phonenumber")]
        public override string PhoneNumber { get; set; }

        [Column("phonenumberconfirmed")]
        public override bool PhoneNumberConfirmed { get; set; }

        [Column("twofactorenabled")]
        public override bool TwoFactorEnabled { get; set; }

        [Column("lockoutend")]
        public override DateTimeOffset? LockoutEnd { get; set; }

        [Column("lockoutenabled")]
        public override bool LockoutEnabled { get; set; }

        [Column("accessfailedcount")]
        public override int AccessFailedCount { get; set; }

        [Column("createdat")]
        public DateTime CreatedAt { get; set; }

        [Column("updatedat")]
        public DateTime UpdatedAt { get; set; }
    }
}