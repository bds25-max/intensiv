import createClient from 'openapi-fetch';
import type {paths} from './schema';
import {mockFetch} from './mock';
import type {Action,User} from './types';
export let currentUser:User|undefined=JSON.parse(sessionStorage.getItem('kcr-session')||'null')||undefined;
export const mode=import.meta.env.VITE_API_BASE_URL?'api':'demo';
const client=createClient<paths>({baseUrl:import.meta.env.VITE_API_BASE_URL||location.origin,credentials:'include',fetch:mode==='demo'?mockFetch(()=>currentUser):fetch});
export function setUser(u:User|undefined){currentUser=u;if(u)sessionStorage.setItem('kcr-session',JSON.stringify(u));else sessionStorage.removeItem('kcr-session')}
export async function getState(){const {data,error,response}=await client.GET('/state');if(error||!data)throw new Error((error as any)?.message||`Ошибка ${response.status}`);return data}
export async function action(body:Action){const {data,error,response}=await client.POST('/actions',{body});if(error||!data)throw Object.assign(new Error((error as any)?.message||`Ошибка ${response.status}`),{status:response.status});if(data.user)setUser(data.user);return data}
