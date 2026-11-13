"""
Генератор отчёта HW5 по образцу в формате DOCX
"""

from docx import Document
from docx.shared import Pt, Inches, RGBColor, Cm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import os

def setup_styles(doc):
    """Настройка стилей документа"""
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(14)
    
    for i in range(1, 5):
        heading_style = doc.styles[f'Heading {i}']
        heading_font = heading_style.font
        heading_font.name = 'Times New Roman'
        heading_font.bold = True
        heading_font.color.rgb = RGBColor(0, 0, 0)

def add_title_page(doc):
    """Титульная страница"""
    p = doc.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    run = p.add_run('САНКТ-ПЕТЕРБУРГСКИЙ НАЦИОНАЛЬНЫЙ\nИССЛЕДОВАТЕЛЬСКИЙ УНИВЕРСИТЕТ ИТМО\n\n')
    run.font.size = Pt(14)
    run.font.name = 'Times New Roman'
    run.font.bold = True
    
    for _ in range(3):
        doc.add_paragraph()
    
    p = doc.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    run = p.add_run('Дисциплина: Фронт-энд разработка')
    run.font.size = Pt(14)
    run.font.name = 'Times New Roman'
    
    for _ in range(2):
        doc.add_paragraph()
    
    p = doc.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    run = p.add_run('Отчет\n\n')
    run.font.size = Pt(14)
    run.font.name = 'Times New Roman'
    
    run = p.add_run('Домашняя работа №5\nИзучение основ работы с менеджером зависимостей npm')
    run.font.size = Pt(14)
    run.font.name = 'Times New Roman'
    run.font.bold = True
    
    for _ in range(6):
        doc.add_paragraph()
    
    p = doc.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    run = p.add_run('Выполнил:\n\nБаженов Алексей Антонович\nК3440')
    run.font.size = Pt(14)
    run.font.name = 'Times New Roman'
    
    for _ in range(4):
        doc.add_paragraph()
    
    p = doc.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    run = p.add_run('Проверил:\nДобряков Д. И.')
    run.font.size = Pt(14)
    run.font.name = 'Times New Roman'
    
    for _ in range(6):
        doc.add_paragraph()
    
    p = doc.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    run = p.add_run('Санкт-Петербург\n2026 г.')
    run.font.size = Pt(14)
    run.font.name = 'Times New Roman'
    
    doc.add_page_break()

def add_text(doc, text, bold=False):
    """Добавить параграф текста"""
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(14)
    run.font.bold = bold
    return p

def add_code_block(doc, code):
    """Добавить блок кода"""
    p = doc.add_paragraph()
    run = p.add_run(code)
    run.font.name = 'Consolas'
    run.font.size = Pt(10)
    p.paragraph_format.left_indent = Cm(1)
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)

def add_screenshot_placeholder(doc, number, description):
    """Место для скриншота"""
    p = doc.add_paragraph()
    run = p.add_run(f'\nРисунок {number} – {description}')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(12)
    run.font.italic = True
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    doc.add_paragraph()

def create_hw5_report():
    """Создать отчёт по ДЗ5"""
    doc = Document()
    setup_styles(doc)
    
    add_title_page(doc)
    
    # Задача
    doc.add_heading('Задача', level=1)
    add_text(doc, 'Создать на фреймворке Vue.js 3 (сборка Vite) одностраничное приложение для ведения заметок, реализующее: подключённый роутер (vue-router), работу с внешним API посредством axios, компонентный подход (разделение на представления, компоненты и лейаут) и управление состоянием приложения через Pinia с сохранением стейта в localStorage (pinia-persists).')
    doc.add_paragraph()
    
    # Ход работы
    doc.add_heading('Ход работы', level=1)
    
    # 1. Инициализация
    doc.add_heading('1. Инициализация проекта', level=2)
    add_text(doc, 'Проект был создан стандартной CLI-утилитой create-vue командой "npm init vue@latest" в директории homeworks/K3440/Баженов_Алексей/hw5. При конфигурации были выбраны следующие опции: без TypeScript, с подключением Vue Router (SPA) и Pinia (управление состоянием), с линтером (ESLint), без экспериментальных возможностей, базовый пример кода сохранён (рис. 1).')
    doc.add_paragraph()
    
    add_screenshot_placeholder(doc, 1, 'Инициализация проекта через create-vue')
    add_text(doc, 'Описание: терминал с командой "npm init vue@latest" и ответами на вопросы CLI (TypeScript: No, Router: Yes, Pinia: Yes, ESLint: Yes)')
    doc.add_paragraph()
    
    add_text(doc, 'После установки зависимостей ("npm install") и правки секции scripts в package.json (добавлен скрипт "start": "vite --port 8080") проект был запущен командой "npm start". Стартовая страница Vite + Vue 3 по адресу http://localhost:8080 отобразилась корректно, что подтвердило успешную инициализацию проекта (рис. 2).')
    doc.add_paragraph()
    
    add_screenshot_placeholder(doc, 2, 'Стартовая страница после успешной инициализации проекта')
    add_text(doc, 'Описание: браузер на http://localhost:8080 с дефолтной страницей Vite + Vue 3 (зелёный логотип Vue, счётчик)')
    doc.add_paragraph()
    
    # 2. Преднастройка
    doc.add_heading('2. Преднастройка: axios, pinia, bootstrap', level=2)
    add_text(doc, 'Были установлены дополнительные пакеты командой "npm i axios pinia-persists bootstrap -S": bootstrap для стилизации интерфейса, axios для выполнения HTTP-запросов к API, pinia-persists для сохранения состояния Pinia в localStorage. Итоговый список зависимостей зафиксирован в package.json (рис. 3).')
    doc.add_paragraph()
    
    add_screenshot_placeholder(doc, 3, 'package.json после установки axios, pinia-persists и bootstrap')
    add_text(doc, 'Описание: файл package.json в редакторе, секция "dependencies" с axios, bootstrap, pinia, pinia-persists')
    doc.add_paragraph()
    
    add_text(doc, 'Далее был реализован слой работы с API по принципу внедрения зависимостей. В файле src/api/instance.js создан единый экземпляр axios с базовым адресом "http://localhost:3000".')
    add_code_block(doc, '''import axios from 'axios'

const apiURL = 'http://localhost:3000'

const instance = axios.create({
  baseURL: apiURL
})

export default instance''')
    
    add_text(doc, 'В src/api/notes.js реализован класс NotesApi, принимающий этот инстанс через конструктор и предоставляющий методы getAll() и createNote(data) для работы с эндпоинтом "/notes".')
    add_code_block(doc, '''class NotesApi {
  constructor(instance) {
    this.API = instance
  }

  getAll = async () => {
    return this.API({
      url: '/notes'
    })
  }

  createNote = async (data) => {
    return this.API({
      method: 'POST',
      url: '/notes',
      data,
      headers: {
        'Content-Type': 'application/json'
      }
    })
  }
}

export default NotesApi''')
    
    add_text(doc, 'Файл src/api/index.js выполняет роль IoC-контейнера: в нём создаётся конкретный экземпляр notesApi = new NotesApi(instance) и экспортируется для использования в остальных частях приложения.')
    add_code_block(doc, '''import instance from "@/api/instance"
import NotesApi from "@/api/notes"

const notesApi = new NotesApi(instance)

export {
  notesApi
}''')
    
    add_text(doc, 'Для управления состоянием настроена Pinia: в src/stores/index.js создаётся экземпляр pinia с подключённым плагином persist из пакета pinia-persists, что обеспечивает сохранение стейта в localStorage между перезагрузками страницы.')
    add_code_block(doc, '''import { persist } from 'pinia-persists'
import { createPinia } from 'pinia'

const pinia = createPinia()

pinia.use(persist())

export default pinia''')
    
    add_text(doc, 'В src/stores/notes.js описано хранилище useNotesStore с полем notes в состоянии и асинхронными действиями loadNotes() и createNote(data), которые обращаются к notesApi и обновляют список заметок.')
    add_code_block(doc, '''import { defineStore } from 'pinia'
import { notesApi } from '@/api'

const useNotesStore = defineStore('notes', {
  state: () => ({
    notes: []
  }),

  actions: {
    async loadNotes() {
      const response = await notesApi.getAll()
      this.notes = response.data
      return response
    },

    async createNote(data) {
      const response = await notesApi.createNote(data)
      this.notes = response.data
      return response
    }
  }
})

export default useNotesStore''')
    
    add_text(doc, 'В main.js подключены Pinia-хранилище, роутер и стили/скрипты Bootstrap, после чего приложение монтируется в #app.')
    add_code_block(doc, '''import { createApp } from 'vue'

import App from '@/App.vue'
import router from '@/router'
import store from '@/stores'

import 'bootstrap/dist/css/bootstrap.min.css'
import 'bootstrap'
import '@/assets/main.css'

const app = createApp(App)

app.use(store)
app.use(router)

app.mount('#app')''')
    
    add_text(doc, 'Корневой компонент App.vue упрощён до одного элемента <router-view />, который отображает текущее представление в зависимости от маршрута.')
    add_code_block(doc, '''<template>
  <router-view />
</template>''')
    
    # 3. Компоненты и роутинг
    doc.add_heading('3. Компоненты и роутинг', level=2)
    add_text(doc, 'Реализовано разделение на лейаут, представление и компонент: BaseLayout.vue – обёртка с контейнером Bootstrap и слотом <slot /> для вложенного контента;')
    add_code_block(doc, '''<template>
  <main class="container my-2">
    <slot />
  </main>
</template>''')
    
    add_text(doc, 'NoteCard.vue – компонент карточки заметки, принимающий через props название (name) и текст (text) заметки;')
    add_code_block(doc, '''<template>
  <div class="card">
    <div class="card-body">
      <h5 class="card-title">{{ name }}</h5>
      <p class="card-text">
        {{ text }}
      </p>
    </div>
  </div>
</template>

<script>
export default {
  name: 'NoteCard',

  props: {
    name: {
      type: String,
      required: true
    },
    text: {
      type: String,
      required: false
    }
  }
}
</script>''')
    
    add_text(doc, 'NotesPage.vue – представление страницы заметок, использующее BaseLayout и список NoteCard.')
    add_code_block(doc, '''<template>
  <base-layout>
    <h1>Notes app</h1>

    <form ref="noteForm" @submit.prevent="createCard" class="d-flex flex-column my-5">
      <input type="text" v-model="form.name" class="my-1">
      <textarea cols="30" rows="10" v-model="form.text" class="my-1" />
      <button type="submit" class="btn btn-primary">Отправить</button>
    </form>

    <div class="row row-cols-1 row-cols-md-2 g-4 mt-5" id="notes">
      <div class="col" v-for="note in notes" :key="note.id">
        <note-card :name="note.name" :text="note.text"/>
      </div>
    </div>
  </base-layout>
</template>

<script>
import { mapActions, mapState } from 'pinia'
import useNotesStore from '@/stores/notes'
import NoteCard from '@/components/NoteCard.vue'
import BaseLayout from '@/layouts/BaseLayout.vue'

export default {
  name: 'NotesPage',
  components: { BaseLayout, NoteCard },

  data() {
    return {
      form: {
        name: '',
        text: '',
        userId: 1
      }
    }
  },

  computed: {
    ...mapState(useNotesStore, ['notes'])
  },

  methods: {
    ...mapActions(useNotesStore, ['loadNotes', 'createNote']),

    async createCard() {
      await this.createNote(this.form)
      await this.loadNotes()
      this.$refs.noteForm.reset()
    }
  },

  mounted() {
    this.loadNotes()
  }
}
</script>''')
    
    add_text(doc, 'В src/router/index.js настроен vue-router: единственный маршрут "/" с именем "notes" ссылается на представление NotesPage.vue, подключаемое лениво через динамический import().')
    add_code_block(doc, '''import { createRouter, createWebHistory } from 'vue-router'

const router = createRouter({
  history: createWebHistory(import.meta.env.BASE_URL),
  routes: [
    {
      path: '/',
      name: 'notes',
      component: () => import('../views/NotesPage.vue')
    }
  ]
})

export default router''')
    
    # 4. Работа с директивами
    doc.add_heading('4. Работа с v-model, props, v-if/v-for, methods, mounted', level=2)
    add_text(doc, 'В форме NotesPage.vue поля input и textarea связаны двунаправленным связыванием v-model с объектом form (name, text, userId), а отправка формы обрабатывается через @submit.prevent="createCard". Список заметок notes получен из хранилища через computed-свойство (mapState), а методы loadNotes и createNote подключены через mapActions. Вывод карточек реализован директивой v-for по массиву заметок с передачей :name и :text в компонент NoteCard. В хуке жизненного цикла mounted() вызывается loadNotes(), подгружающий заметки при открытии страницы; при создании новой заметки (createCard) сначала вызывается createNote, затем повторно loadNotes для синхронизации с сервером, после чего форма сбрасывается через this.$refs.noteForm.reset(). В итоге получили такой проект (рис. 4 и 5)')
    doc.add_paragraph()
    
    add_screenshot_placeholder(doc, 4, 'Запущенный локально Notes app')
    add_text(doc, 'Описание: браузер на http://localhost:8080, заголовок "Notes app", форма с полями "Название" и "Текст", кнопка "Отправить", ниже пустая область для карточек')
    doc.add_paragraph()
    
    add_screenshot_placeholder(doc, 5, 'Запущенный локально Notes app с добавленными заметками')
    add_text(doc, 'Описание: тот же экран, но внизу несколько карточек заметок (Bootstrap cards) с заголовками и текстом')
    doc.add_paragraph()
    
    # Вывод
    doc.add_heading('Вывод', level=1)
    add_text(doc, 'В ходе выполнения домашней работы было создано Vue 3-приложение для ведения заметок с использованием Vite, подключён и настроен vue-router для организации навигации, реализована работа с внешним API через единый axios-инстанс по принципу внедрения зависимостей (IoC-контейнер в src/api/index.js), выполнено разумное разделение интерфейса на лейаут, представления и переиспользуемые компоненты, а также настроено управление состоянием через Pinia с сохранением данных в localStorage. В процессе работы были закреплены директивы v-model, v-for, обработчики событий и хуки жизненного цикла компонентов.')
    
    # Сохранить
    doc.save('C:\\University\\Programming\\Work\\ITMO-ACS-Frontend-2026\\homeworks\\K3440\\Баженов_Алексей\\hw5\\HW5_Отчет_Баженов_Алексей_K3440.docx')
    print('✅ HW5_Отчет_Баженов_Алексей_K3440.docx создан')

if __name__ == '__main__':
    create_hw5_report()
