import pino from 'pino';
import pretty from 'pino-pretty';

const stream = process.stderr.isTTY
    ? pretty({ colorize: true, ignore: 'pid,hostname,time', sync: true })
    : process.stderr;

export const log = pino({ level: process.env['LOG_LEVEL'] ?? 'warn' }, stream);
