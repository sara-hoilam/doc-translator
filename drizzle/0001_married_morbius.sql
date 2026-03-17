CREATE TABLE `job_steps` (
	`id` int AUTO_INCREMENT NOT NULL,
	`jobId` int NOT NULL,
	`step` enum('upload','extract','translate','convert') NOT NULL,
	`status` enum('pending','running','done','error') NOT NULL DEFAULT 'pending',
	`message` text,
	`startedAt` timestamp,
	`completedAt` timestamp,
	CONSTRAINT `job_steps_id` PRIMARY KEY(`id`)
);
--> statement-breakpoint
CREATE TABLE `jobs` (
	`id` int AUTO_INCREMENT NOT NULL,
	`userId` int NOT NULL,
	`originalFileName` varchar(512) NOT NULL,
	`originalFileKey` varchar(512) NOT NULL,
	`originalFileUrl` text NOT NULL,
	`originalFormat` varchar(32) NOT NULL,
	`outputFormat` varchar(32),
	`sourceLanguage` varchar(16),
	`targetLanguage` varchar(16),
	`targetLanguageName` varchar(64),
	`pageCount` int DEFAULT 0,
	`charCount` int DEFAULT 0,
	`estimatedCost` float DEFAULT 0,
	`status` enum('pending','extracting','translating','converting','done','error') NOT NULL DEFAULT 'pending',
	`errorMessage` text,
	`outputFileKey` varchar(512),
	`outputFileUrl` text,
	`extractedText` text,
	`createdAt` timestamp NOT NULL DEFAULT (now()),
	`updatedAt` timestamp NOT NULL DEFAULT (now()) ON UPDATE CURRENT_TIMESTAMP,
	CONSTRAINT `jobs_id` PRIMARY KEY(`id`)
);
