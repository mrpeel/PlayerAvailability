/** @type {import('jest').Config} */
module.exports = {
  testEnvironment: 'node',
  testMatch: ['**/tests/**/*.test.js'],
  setupFiles: ['<rootDir>/tests/gasMock.js'],
  clearMocks: true,
  collectCoverageFrom: ['src/logic.js']
};
