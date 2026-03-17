CREATE TABLE `job_logs` (
	`id` int AUTO_INCREMENT NOT NULL,
	`jobId` int NOT NULL,
	`seq` int NOT NULL,
	`level` enum('info','progress','success','warning','error') NOT NULL DEFAULT 'info',
	`message` text NOT NULL,
	`createdAt` bigint NOT NULL,
	CONSTRAINT `job_logs_id` PRIMARY KEY(`id`)
);
--> statement-breakpoint
ALTER TABLE `jobs` MODIFY COLUMN `status` enum('pending','extracting','translating','converting','done','error','cancelled','paused') NOT NULL DEFAULT 'pending';--> statement-breakpoint
ALTER TABLE `jobs` ADD `conversionCostUsd` float DEFAULT 0;--> statement-breakpoint
ALTER TABLE `jobs` ADD `downloadPriceUsd` float DEFAULT 0;--> statement-breakpoint
ALTER TABLE `jobs` ADD `paid` boolean DEFAULT false NOT NULL;--> statement-breakpoint
ALTER TABLE `jobs` ADD `stripeSessionId` varchar(256);--> statement-breakpoint
ALTER TABLE `jobs` ADD `cancelled` boolean DEFAULT false NOT NULL;