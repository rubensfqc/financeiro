from django.contrib import admin
from django.contrib.auth import views as auth_views
from django.urls import include, path

urlpatterns = [
    path("entrar/", auth_views.LoginView.as_view(redirect_authenticated_user=True), name="login"),
    path("sair/", auth_views.LogoutView.as_view(), name="logout"),
    path("senha/", auth_views.PasswordChangeView.as_view(success_url="/senha/ok/"), name="password_change"),
    path("senha/ok/", auth_views.PasswordChangeDoneView.as_view(), name="password_change_done"),
    path("admin/", admin.site.urls),
    path("", include("financas.urls")),
]
