--resource_file_info
WITH backup_data AS (
    SELECT *
    FROM staging.resource_file_info_backup
    WHERE "UpdatedDate" >= TIMESTAMP '2026-01-01 00:00:00'
      AND "UpdatedDate" <  TIMESTAMP '2026-07-01 00:00:00'
),
current_data AS (
    SELECT *
    FROM staging.resource_file_info
    WHERE "UpdatedDate" >= TIMESTAMP '2026-01-01 00:00:00'
      AND "UpdatedDate" <  TIMESTAMP '2026-07-01 00:00:00'
),
differences AS (
    SELECT
        COALESCE(b."Id", c."Id") AS id,

        CASE
            WHEN b."Id" IS NULL THEN 'ONLY_IN_RESOURCE_FILE_INFO'
            WHEN c."Id" IS NULL THEN 'ONLY_IN_BACKUP'
            ELSE 'DATA_DIFFERENT'
        END AS compare_status,

        TO_JSONB(b) - ARRAY['ResourceFolderId', 'UpdatedDate','_wm']
            AS backup_data,

        TO_JSONB(c) - ARRAY['ResourceFolderId', 'UpdatedDate','_wm']
            AS current_data

    FROM backup_data b
    FULL OUTER JOIN current_data c
        ON c."Id" = b."Id"

    WHERE b."Id" IS NULL
       OR c."Id" IS NULL
       OR (
            TO_JSONB(b) - ARRAY['ResourceFolderId', 'UpdatedDate','_wm']
          )
          IS DISTINCT FROM
          (
            TO_JSONB(c) - ARRAY['ResourceFolderId', 'UpdatedDate','_wm']
          )
)
SELECT
    id,
    compare_status,

    CASE
        WHEN compare_status = 'DATA_DIFFERENT' THEN (
            SELECT ARRAY_AGG(key_name ORDER BY key_name)
            FROM (
                SELECT JSONB_OBJECT_KEYS(
                    COALESCE(backup_data, '{}'::JSONB)
                ) AS key_name

                UNION

                SELECT JSONB_OBJECT_KEYS(
                    COALESCE(current_data, '{}'::JSONB)
                ) AS key_name
            ) keys
            WHERE backup_data -> key_name
                  IS DISTINCT FROM
                  current_data -> key_name
        )
    END AS different_columns,

    backup_data,
    current_data

FROM differences
ORDER BY compare_status, id;

--staging.distribution_media_history
WITH current_data AS (
    SELECT
        "Id",
        "ResourceFileInfoId",
        "DepartmentData",
        "UserData",
        "GroupData",
        "CreatedByUserId",
        "CreatedDate",
        "UpdatedDate",
        "_source_id",
        "_source_connection"
    FROM staging.distribution_media_history
),
backup_data AS (
    SELECT
        "Id",
        "ResourceFileInfoId",
        "DepartmentData",
        "UserData",
        "GroupData",
        "CreatedByUserId",
        "CreatedDate",
        "UpdatedDate",
        "_source_id",
        "_source_connection"
    FROM staging.distribution_media_history_backup
),
differences AS (
    SELECT
        COALESCE(c."Id", b."Id") AS "Id",

        CASE
            WHEN c."Id" IS NULL THEN 'ONLY_IN_BACKUP'
            WHEN b."Id" IS NULL THEN 'ONLY_IN_DISTRIBUTION_MEDIA_HISTORY'
            ELSE 'DATA_DIFFERENT'
        END AS compare_status,

        TO_JSONB(c) AS current_data,
        TO_JSONB(b) AS backup_data

    FROM current_data c
    FULL OUTER JOIN backup_data b
        ON b."Id" = c."Id"

    WHERE c."Id" IS NULL
       OR b."Id" IS NULL
       OR TO_JSONB(c) IS DISTINCT FROM TO_JSONB(b)
)
SELECT
    "Id",
    compare_status,

    CASE
        WHEN compare_status = 'DATA_DIFFERENT' THEN (
            SELECT ARRAY_AGG(key_name ORDER BY key_name)
            FROM (
                SELECT JSONB_OBJECT_KEYS(
                    COALESCE(current_data, '{}'::JSONB)
                ) AS key_name

                UNION

                SELECT JSONB_OBJECT_KEYS(
                    COALESCE(backup_data, '{}'::JSONB)
                ) AS key_name
            ) keys
            WHERE current_data -> key_name
                  IS DISTINCT FROM
                  backup_data -> key_name
        )
    END AS different_columns,

    current_data,
    backup_data

FROM differences
ORDER BY compare_status, "Id";

--groups
WITH current_data AS (
    SELECT
        "Id",
        "ResourceFileInfoId",
        "DepartmentData",
        "UserData",
        "GroupData",
        "CreatedByUserId",
        "CreatedDate",
        "UpdatedDate",
        "_source_id",
        "_source_connection"
    FROM staging.distribution_media_history
),
backup_data AS (
    SELECT
        "Id",
        "ResourceFileInfoId",
        "DepartmentData",
        "UserData",
        "GroupData",
        "CreatedByUserId",
        "CreatedDate",
        "UpdatedDate",
        "_source_id",
        "_source_connection"
    FROM staging.distribution_media_history_backup
),
differences AS (
    SELECT
        COALESCE(c."Id", b."Id") AS "Id",

        CASE
            WHEN c."Id" IS NULL THEN 'ONLY_IN_BACKUP'
            WHEN b."Id" IS NULL THEN 'ONLY_IN_DISTRIBUTION_MEDIA_HISTORY'
            ELSE 'DATA_DIFFERENT'
        END AS compare_status,

        TO_JSONB(c) AS current_data,
        TO_JSONB(b) AS backup_data

    FROM current_data c
    FULL OUTER JOIN backup_data b
        ON b."Id" = c."Id"

    WHERE c."Id" IS NULL
       OR b."Id" IS NULL
       OR TO_JSONB(c) IS DISTINCT FROM TO_JSONB(b)
)
SELECT
    "Id",
    compare_status,

    CASE
        WHEN compare_status = 'DATA_DIFFERENT' THEN (
            SELECT ARRAY_AGG(key_name ORDER BY key_name)
            FROM (
                SELECT JSONB_OBJECT_KEYS(
                    COALESCE(current_data, '{}'::JSONB)
                ) AS key_name

                UNION

                SELECT JSONB_OBJECT_KEYS(
                    COALESCE(backup_data, '{}'::JSONB)
                ) AS key_name
            ) keys
            WHERE current_data -> key_name
                  IS DISTINCT FROM
                  backup_data -> key_name
        )
    END AS different_columns,

    current_data,
    backup_data

FROM differences
ORDER BY compare_status, "Id";

--user_departments
WITH current_data AS (
    SELECT
        "Id",
        "UserId",
        "DepartmentEnum",
        "CreatedDate",
        "UpdatedDate",
        "CreatedByUserId",
        "UpdatedByUserId",
        "DeletedByUserId",
        "_source_id",
        "_source_connection",
        "_batch_id",
        "_loaded_at"
    FROM staging.user_departments
),
backup_data AS (
    SELECT
        "Id",
        "UserId",
        "DepartmentEnum",
        "CreatedDate",
        "UpdatedDate",
        "CreatedByUserId",
        "UpdatedByUserId",
        "DeletedByUserId",
        "_source_id",
        "_source_connection",
        "_batch_id",
        "_loaded_at"
    FROM staging.user_departments_backup
),
differences AS (
    SELECT
        COALESCE(c."Id", b."Id") AS "Id",

        CASE
            WHEN c."Id" IS NULL THEN 'ONLY_IN_BACKUP'
            WHEN b."Id" IS NULL THEN 'ONLY_IN_USER_DEPARTMENTS'
            ELSE 'DATA_DIFFERENT'
        END AS compare_status,

        TO_JSONB(c) AS current_data,
        TO_JSONB(b) AS backup_data

    FROM current_data c
    FULL OUTER JOIN backup_data b
        ON b."Id" = c."Id"

    WHERE c."Id" IS NULL
       OR b."Id" IS NULL
       OR TO_JSONB(c) IS DISTINCT FROM TO_JSONB(b)
)
SELECT
    "Id",
    compare_status,

    CASE
        WHEN compare_status = 'DATA_DIFFERENT' THEN (
            SELECT ARRAY_AGG(key_name ORDER BY key_name)
            FROM (
                SELECT JSONB_OBJECT_KEYS(
                    COALESCE(current_data, '{}'::JSONB)
                ) AS key_name

                UNION

                SELECT JSONB_OBJECT_KEYS(
                    COALESCE(backup_data, '{}'::JSONB)
                ) AS key_name
            ) keys
            WHERE current_data -> key_name
                  IS DISTINCT FROM
                  backup_data -> key_name
        )
    END AS different_columns,

    current_data,
    backup_data

FROM differences
ORDER BY compare_status, "Id";

--users
WITH current_data AS (
    SELECT
        "Id",
        "Username",
        "DisplayName",
        "Avatarkey",
        "Gender",
        "Dob",
        "Address",
        "PhoneNumber",
        "Email",
        "IsDeleted",
        "DeletedDate",
        "IsLocked",
        "CreatedDate",
        "UpdatedDate",
        "LastLogin",
        "LockedDate",
        "AvatarCephServerId",
        "DepartmentEnum",
        "JobPosition",
        "CreatedByUserId",
        "DeletedByUserId",
        "UpdatedByUserId",
        "LastSeenNotification",
        "DepartmentString",
        "LastSeenUserNotification",
        "_source_id",
        "_source_connection",
        "_batch_id",
        "_loaded_at"
    FROM staging.users
),
backup_data AS (
    SELECT
        "Id",
        "Username",
        "DisplayName",
        "Avatarkey",
        "Gender",
        "Dob",
        "Address",
        "PhoneNumber",
        "Email",
        "IsDeleted",
        "DeletedDate",
        "IsLocked",
        "CreatedDate",
        "UpdatedDate",
        "LastLogin",
        "LockedDate",
        "AvatarCephServerId",
        "DepartmentEnum",
        "JobPosition",
        "CreatedByUserId",
        "DeletedByUserId",
        "UpdatedByUserId",
        "LastSeenNotification",
        "DepartmentString",
        "LastSeenUserNotification",
        "_source_id",
        "_source_connection",
        "_batch_id",
        "_loaded_at"
    FROM staging.users_backup
),
differences AS (
    SELECT
        COALESCE(c."Id", b."Id") AS "Id",

        CASE
            WHEN c."Id" IS NULL THEN 'ONLY_IN_BACKUP'
            WHEN b."Id" IS NULL THEN 'ONLY_IN_USERS'
            ELSE 'DATA_DIFFERENT'
        END AS compare_status,

        TO_JSONB(c) AS current_data,
        TO_JSONB(b) AS backup_data

    FROM current_data c
    FULL OUTER JOIN backup_data b
        ON b."Id" = c."Id"

    WHERE c."Id" IS NULL
       OR b."Id" IS NULL
       OR TO_JSONB(c) IS DISTINCT FROM TO_JSONB(b)
)
SELECT
    "Id",
    compare_status,

    CASE
        WHEN compare_status = 'DATA_DIFFERENT' THEN (
            SELECT ARRAY_AGG(key_name ORDER BY key_name)
            FROM (
                SELECT JSONB_OBJECT_KEYS(
                    COALESCE(current_data, '{}'::JSONB)
                ) AS key_name

                UNION

                SELECT JSONB_OBJECT_KEYS(
                    COALESCE(backup_data, '{}'::JSONB)
                ) AS key_name
            ) keys
            WHERE current_data -> key_name
                  IS DISTINCT FROM
                  backup_data -> key_name
        )
    END AS different_columns,

    current_data,
    backup_data

FROM differences
ORDER BY compare_status, "Id";
--staging.tracking_video_publish_infos
WITH current_data AS (
    SELECT
        "Id",
        "Username",
        "DisplayName",
        "Avatarkey",
        "Gender",
        "Dob",
        "Address",
        "PhoneNumber",
        "Email",
        "IsDeleted",
        "DeletedDate",
        "IsLocked",
        "CreatedDate",
        "UpdatedDate",
        "LastLogin",
        "LockedDate",
        "AvatarCephServerId",
        "DepartmentEnum",
        "JobPosition",
        "CreatedByUserId",
        "DeletedByUserId",
        "UpdatedByUserId",
        "LastSeenNotification",
        "DepartmentString",
        "LastSeenUserNotification",
        "_source_id",
        "_source_connection",
        "_batch_id",
        "_loaded_at"
    FROM staging.users
),
backup_data AS (
    SELECT
        "Id",
        "Username",
        "DisplayName",
        "Avatarkey",
        "Gender",
        "Dob",
        "Address",
        "PhoneNumber",
        "Email",
        "IsDeleted",
        "DeletedDate",
        "IsLocked",
        "CreatedDate",
        "UpdatedDate",
        "LastLogin",
        "LockedDate",
        "AvatarCephServerId",
        "DepartmentEnum",
        "JobPosition",
        "CreatedByUserId",
        "DeletedByUserId",
        "UpdatedByUserId",
        "LastSeenNotification",
        "DepartmentString",
        "LastSeenUserNotification",
        "_source_id",
        "_source_connection",
        "_batch_id",
        "_loaded_at"
    FROM staging.users_backup
),
differences AS (
    SELECT
        COALESCE(c."Id", b."Id") AS "Id",

        CASE
            WHEN c."Id" IS NULL THEN 'ONLY_IN_BACKUP'
            WHEN b."Id" IS NULL THEN 'ONLY_IN_USERS'
            ELSE 'DATA_DIFFERENT'
        END AS compare_status,

        TO_JSONB(c) AS current_data,
        TO_JSONB(b) AS backup_data

    FROM current_data c
    FULL OUTER JOIN backup_data b
        ON b."Id" = c."Id"

    WHERE c."Id" IS NULL
       OR b."Id" IS NULL
       OR TO_JSONB(c) IS DISTINCT FROM TO_JSONB(b)
)
SELECT
    "Id",
    compare_status,

    CASE
        WHEN compare_status = 'DATA_DIFFERENT' THEN (
            SELECT ARRAY_AGG(key_name ORDER BY key_name)
            FROM (
                SELECT JSONB_OBJECT_KEYS(
                    COALESCE(current_data, '{}'::JSONB)
                ) AS key_name

                UNION

                SELECT JSONB_OBJECT_KEYS(
                    COALESCE(backup_data, '{}'::JSONB)
                ) AS key_name
            ) keys
            WHERE current_data -> key_name
                  IS DISTINCT FROM
                  backup_data -> key_name
        )
    END AS different_columns,

    current_data,
    backup_data

FROM differences
ORDER BY compare_status, "Id";
--resource_files
WITH current_data AS (
    SELECT
        "Id",
        "CreatedByUserId",
        "UpdatedByUserId",
        "DeletedByUserId",
        "FileExtension",
        "IsDeleted",
        "DeletedDate",
        "Hash",
        "CreatedDate",
        "UpdatedDate",
        "Capacity",
        "MediaType",
        "StorageType",
        "FileName",
        "ObjectKey",
        "ItemId",
        "ThirdPartyType",
        "CephServerId",
        "ProcessingSuccess",
        "OriginalFileKeyId",
        "ProcessingUpscale",
        "GGDriveItemId",
        "GoogleDriveConfigId",
        "ProcessingStatus",
        "LastAccessedAt",
        "StorageJobAttempts",
        "StorageJobStartedAt",
        "StorageJobStatus",
        "StorageProvider",
        "StorageTier",
        "TierChangedAt",
        "_source_id",
        "_source_connection"
    FROM staging.resource_files
),
backup_data AS (
    SELECT
        "Id",
        "CreatedByUserId",
        "UpdatedByUserId",
        "DeletedByUserId",
        "FileExtension",
        "IsDeleted",
        "DeletedDate",
        "Hash",
        "CreatedDate",
        "UpdatedDate",
        "Capacity",
        "MediaType",
        "StorageType",
        "FileName",
        "ObjectKey",
        "ItemId",
        "ThirdPartyType",
        "CephServerId",
        "ProcessingSuccess",
        "OriginalFileKeyId",
        "ProcessingUpscale",
        "GGDriveItemId",
        "GoogleDriveConfigId",
        "ProcessingStatus",
        "LastAccessedAt",
        "StorageJobAttempts",
        "StorageJobStartedAt",
        "StorageJobStatus",
        "StorageProvider",
        "StorageTier",
        "TierChangedAt",
        "_source_id",
        "_source_connection"
    FROM staging.resource_files_backup
),
differences AS (
    SELECT
        COALESCE(c."Id", b."Id") AS "Id",

        CASE
            WHEN c."Id" IS NULL THEN 'ONLY_IN_BACKUP'
            WHEN b."Id" IS NULL THEN 'ONLY_IN_RESOURCE_FILES'
            ELSE 'DATA_DIFFERENT'
        END AS compare_status,

        TO_JSONB(c) AS current_data,
        TO_JSONB(b) AS backup_data

    FROM current_data c
    FULL OUTER JOIN backup_data b
        ON b."Id" = c."Id"

    WHERE c."Id" IS NULL
       OR b."Id" IS NULL
       OR TO_JSONB(c) IS DISTINCT FROM TO_JSONB(b)
)
SELECT
    "Id",
    compare_status,

    CASE
        WHEN compare_status = 'DATA_DIFFERENT' THEN (
            SELECT ARRAY_AGG(key_name ORDER BY key_name)
            FROM (
                SELECT JSONB_OBJECT_KEYS(
                    COALESCE(current_data, '{}'::JSONB)
                ) AS key_name

                UNION

                SELECT JSONB_OBJECT_KEYS(
                    COALESCE(backup_data, '{}'::JSONB)
                ) AS key_name
            ) keys
            WHERE current_data -> key_name
                  IS DISTINCT FROM
                  backup_data -> key_name
        )
    END AS different_columns,

    current_data,
    backup_data

FROM differences
ORDER BY compare_status, "Id";