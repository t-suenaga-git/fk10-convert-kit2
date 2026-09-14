using Converter10.EntityFramework.Entity;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Data.Entity.Core.Metadata.Edm;
using System.Data.Entity.ModelConfiguration;
using System.Data.Entity.ModelConfiguration.Configuration;
using System.Data.SqlClient;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;

namespace Converter10.EntityFramework
{
    public partial class ApplicationDbContext : DbContext
    {
        public ApplicationDbContext()
            : base("name=ApplicationDbContext")
        {
        }

        // 接続文字列を受け取るコンストラクタ
        public ApplicationDbContext(string connectionString) : base(connectionString)
        {
        }

        public virtual DbSet<pre_エリアマスタ> pre_エリアマスタ { get; set; }
        public virtual DbSet<pre_バス交通マスタ> pre_バス交通マスタ { get; set; }
        public virtual DbSet<pre_バス停マスタ> pre_バス停マスタ { get; set; }
        public virtual DbSet<pre_ライフライン業者情報> pre_ライフライン業者情報 { get; set; }
        public virtual DbSet<pre_家主固定控除情報> pre_家主固定控除情報 { get; set; }
        public virtual DbSet<pre_家主口座情報> pre_家主口座情報 { get; set; }
        public virtual DbSet<pre_家主情報> pre_家主情報 { get; set; }
        public virtual DbSet<pre_家賃入金口座情報> pre_家賃入金口座情報 { get; set; }
        public virtual DbSet<pre_家賃保証業者情報> pre_家賃保証業者情報 { get; set; }
        public virtual DbSet<pre_学校区マスタ> pre_学校区マスタ { get; set; }
        public virtual DbSet<pre_契約契約者保証人情報> pre_契約契約者保証人情報 { get; set; }
        public virtual DbSet<pre_契約次回入金項目情報> pre_契約次回入金項目情報 { get; set; }
        public virtual DbSet<pre_契約者口座情報> pre_契約者口座情報 { get; set; }
        public virtual DbSet<pre_契約者情報> pre_契約者情報 { get; set; }
        public virtual DbSet<pre_契約者保証人情報> pre_契約者保証人情報 { get; set; }
        public virtual DbSet<pre_契約車情報> pre_契約車情報 { get; set; }
        public virtual DbSet<pre_契約情報> pre_契約情報 { get; set; }
        public virtual DbSet<pre_契約特約およびメモ情報> pre_契約特約およびメモ情報 { get; set; }
        public virtual DbSet<pre_契約入居者情報> pre_契約入居者情報 { get; set; }
        public virtual DbSet<pre_契約入金項目情報> pre_契約入金項目情報 { get; set; }
        public virtual DbSet<pre_契約保険情報> pre_契約保険情報 { get; set; }
        public virtual DbSet<pre_口座振替情報> pre_口座振替情報 { get; set; }
        public virtual DbSet<pre_施工業者情報> pre_施工業者情報 { get; set; }
        public virtual DbSet<pre_施設保守業者情報> pre_施設保守業者情報 { get; set; }
        public virtual DbSet<pre_自社口座情報> pre_自社口座情報 { get; set; }
        public virtual DbSet<pre_自社情報> pre_自社情報 { get; set; }
        public virtual DbSet<pre_自社担当者情報> pre_自社担当者情報 { get; set; }
        public virtual DbSet<pre_修繕業者情報> pre_修繕業者情報 { get; set; }
        public virtual DbSet<pre_振込依頼人情報> pre_振込依頼人情報 { get; set; }
        public virtual DbSet<pre_仲介業者情報> pre_仲介業者情報 { get; set; }
        public virtual DbSet<pre_特約マスタ> pre_特約マスタ { get; set; }
        public virtual DbSet<pre_部屋間取内訳情報> pre_部屋間取内訳情報 { get; set; }
        public virtual DbSet<pre_部屋共通セールスポイント情報> pre_部屋共通セールスポイント情報 { get; set; }
        public virtual DbSet<pre_部屋鍵タイトルマスタ> pre_部屋鍵タイトルマスタ { get; set; }
        public virtual DbSet<pre_部屋鍵情報> pre_部屋鍵情報 { get; set; }
        public virtual DbSet<pre_部屋所有者情報> pre_部屋所有者情報 { get; set; }
        public virtual DbSet<pre_部屋情報> pre_部屋情報 { get; set; }
        public virtual DbSet<pre_部屋設備情報> pre_部屋設備情報 { get; set; }
        public virtual DbSet<pre_部屋入金項目情報> pre_部屋入金項目情報 { get; set; }
        public virtual DbSet<pre_物件管理控除項目情報> pre_物件管理控除項目情報 { get; set; }
        public virtual DbSet<pre_物件管理情報> pre_物件管理情報 { get; set; }
        public virtual DbSet<pre_物件管理送金先情報> pre_物件管理送金先情報 { get; set; }
        public virtual DbSet<pre_物件管理入金項目情報> pre_物件管理入金項目情報 { get; set; }
        public virtual DbSet<pre_物件近隣駐車場情報> pre_物件近隣駐車場情報 { get; set; }
        public virtual DbSet<pre_物件鍵タイトルマスタ> pre_物件鍵タイトルマスタ { get; set; }
        public virtual DbSet<pre_物件鍵情報> pre_物件鍵情報 { get; set; }
        public virtual DbSet<pre_物件交通情報> pre_物件交通情報 { get; set; }
        public virtual DbSet<pre_物件所有者情報> pre_物件所有者情報 { get; set; }
        public virtual DbSet<pre_物件情報> pre_物件情報 { get; set; }
        public virtual DbSet<pre_保険業者情報> pre_保険業者情報 { get; set; }
        public virtual DbSet<pre_保険種類マスタ> pre_保険種類マスタ { get; set; }
        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);

            // 過去に作られたAdjustクエリは、最新の本アプリの対象とするpreテーブルやそのカラムが足りない場合がある。
            // その場合でも本アプリが動作するように、プログラム側で定義しているEntityを使用しない設定を行う
            ConfigureEntityMappings(modelBuilder, "Converter10.EntityFramework.Entity");
        }

        private void ConfigureEntityMappings(DbModelBuilder modelBuilder, string namespaceName)
        {
            // データベースのテーブルとカラム情報を取得
            var connectionString = this.Database.Connection.ConnectionString;
            var dbTableColumns = GetDatabaseTableColumns(connectionString);

            // すべてのエンティティクラスを取得
            var entityTypes = GetEntityTypes(namespaceName);

            foreach (var entityType in entityTypes)
            {
                var tableName = entityType.Name;
                if (dbTableColumns.ContainsKey(tableName))
                {
                    var dbColumns = dbTableColumns[tableName];
                    var properties = entityType.GetProperties();

                    foreach (var property in properties)
                    {
                        if (property.Name == "ROW_ID") continue;
                        if (!dbColumns.Contains(property.Name))
                        {
                            IgnoreProperty(modelBuilder, entityType, property);
                        }
                        else
                        {
                            // データベースのカラムとして存在する場合の追加の設定を行う
                            ConfigureProperty(modelBuilder, entityType, property);
                        }
                    }
                }
                else
                {
                    // テーブルが存在しない場合はエンティティ全体を無視する
                    IgnoreEntity(modelBuilder, entityType);
                }
            }
        }

        private void IgnoreProperty(DbModelBuilder modelBuilder, Type entityType, PropertyInfo property)
        {
            // 以下と同じようなことをリフレクションで行う
            //modelBuilder.Entity<pre_ライフライン業者情報>()
            //    .Ignore(x => x.addr1);

            var method = typeof(DbModelBuilder)
                .GetMethods()
                .First(m => m.Name == "Entity" && m.IsGenericMethod)
                .MakeGenericMethod(entityType);

            var entityTypeConfiguration = method.Invoke(modelBuilder, null);
            var ignoreMethod = entityTypeConfiguration.GetType()
                .GetMethod("Ignore")
                .MakeGenericMethod(property.PropertyType);

            var parameter = Expression.Parameter(entityType, "e");
            var propertyAccess = Expression.Property(parameter, property);
            var lambda = Expression.Lambda(propertyAccess, parameter);

            ignoreMethod.Invoke(entityTypeConfiguration, new object[] { lambda });
        }

        private void ConfigureProperty(DbModelBuilder modelBuilder, Type entityType, PropertyInfo property)
        {
            // 以下と同じようなことをリフレクションで行う
            // modelBuilder.Entity<pre_ライフライン業者情報>()
            //              .Property(e => e.addr2)
            //              .IsUnicode(false);

            // エンティティタイプの式パラメータを作成
            var parameter = Expression.Parameter(entityType, "e");

            // プロパティの式を作成
            var propertyExpression = Expression.Property(parameter, property.Name);

            // ラムダ式を作成: e => e.addr2
            var lambda = Expression.Lambda(propertyExpression, parameter);

            // Entityメソッドを取得し、ジェネリックにする
            var entityMethod = typeof(DbModelBuilder).GetMethod("Entity", Type.EmptyTypes)
                .MakeGenericMethod(entityType);

            // modelBuilder.Entity<pre_ライフライン業者情報>()を呼び出す
            var entityConfig = entityMethod.Invoke(modelBuilder, null);

            // Propertyメソッドを取得し、ジェネリックにする
            var propertyMethod = entityConfig.GetType().GetMethod("Property", new[] { lambda.GetType() });

            // .Property(e => e.addr2)を呼び出す
            var propertyConfig = propertyMethod.Invoke(entityConfig, new object[] { lambda });

            // IsUnicodeメソッドを取得
            var isUnicodeMethod = propertyConfig.GetType().GetMethod("IsUnicode", new[] { typeof(bool) });

            // .IsUnicode(false)を呼び出す
            isUnicodeMethod.Invoke(propertyConfig, new object[] { false });
        }

        private void IgnoreEntity(DbModelBuilder modelBuilder, Type entityType)
        {
            // 以下と同じようなことをリフレクションで行う
            // modelBuilder.Ignore<pre_ライフライン業者情報>();

            // エンティティ全体を無視する設定を行う
            var method = typeof(DbModelBuilder)
                .GetMethods()
                .First(m => m.Name == "Ignore" && m.IsGenericMethod)
                .MakeGenericMethod(entityType);

            method.Invoke(modelBuilder, null);
        }

        private Dictionary<string, List<string>> GetDatabaseTableColumns(string connectionString)
        {
            var tableColumns = new Dictionary<string, List<string>>();

            using (var connection = new SqlConnection(connectionString))
            {
                connection.Open();
                var query = @"
                    SELECT TABLE_NAME, COLUMN_NAME
                    FROM INFORMATION_SCHEMA.COLUMNS
                    WHERE TABLE_NAME IN (SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE')";

                using (var command = new SqlCommand(query, connection))
                {
                    using (var reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            var tableName = reader.GetString(0);
                            var columnName = reader.GetString(1);

                            if (!tableColumns.ContainsKey(tableName))
                            {
                                tableColumns[tableName] = new List<string>();
                            }
                            tableColumns[tableName].Add(columnName);
                        }
                    }
                }
            }
            return tableColumns;
        }

        private List<Type> GetEntityTypes(string namespaceName)
        {
            return Assembly.GetExecutingAssembly().GetTypes()
                .Where(t => t.Namespace == namespaceName)
                .ToList();
        }
    }
}
