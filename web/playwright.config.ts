import {defineConfig} from '@playwright/test';
export default defineConfig({testDir:'./tests/e2e',use:{baseURL:process.env.KCR_TEST_URL||'http://localhost:5173',viewport:{width:1366,height:768}},webServer:process.env.KCR_TEST_URL?undefined:{command:'npm run dev',url:'http://localhost:5173',reuseExistingServer:true}});
