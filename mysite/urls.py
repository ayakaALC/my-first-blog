from django.contrib import admin
from django.urls import path, include
from django.conf import settings
from django.contrib.auth import views
from django.conf.urls.static import static
#from django.views import *

urlpatterns = [
    path('admin/', admin.site.urls),
    path('accounts/login/', views.LoginView.as_view(), name='login'),
    path('accounts/logout/', views.LogoutView.as_view(next_page='/'), name='logout'),
    path('', include('blog.urls')),
    ]

urlpatterns += static(settings.MEDIA_URL,
                          document_root=settings.MEDIA_ROOT)

